# !/usr/bin/python
# -*- coding: UTF-8 -*-
##################################################
# __Date__: 2024/5/30
# __Python__: 3.9.6
# __Author__: Jianfan.Ai
##################################################

import time
import datetime
import logging
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from rich.logging import RichHandler
from rich.progress import track
from Phidget22.Phidget import *
from Phidget22.Devices.LightSensor import *
from Phidget22.Devices.SoundSensor import *
from collections import defaultdict
import numpy as np

# Suppress font-related debug messages from matplotlib
matplotlib_logger = logging.getLogger('matplotlib')
matplotlib_logger.setLevel(logging.WARNING)

class VintHubController:
    def __init__(self, light_sensor_port=None, sound_sensor_port=None, hub_serial_number=None, verbose=False):
        """
        Initialize the VintHubController with specified ports and serial number.

        :param light_sensor_port: The hub port for the LightSensor.
        :param sound_sensor_port: The hub port for the SoundSensor.
        :param hub_serial_number: The serial number for the VINT Hub.
        :param verbose: A flag for verbose logging.
        """
        self.verbose = verbose
        self.init_logging()

        self.light_sensor = LightSensor()
        self.sound_sensor = SoundSensor()

        # Set Hub Port
        if light_sensor_port is not None:
            self.light_sensor.setHubPort(light_sensor_port)  # LightSensor Port
        if sound_sensor_port is not None:
            self.sound_sensor.setHubPort(sound_sensor_port)  # SoundSensor Port

        # Optionally set Hub serial number if provided
        if hub_serial_number:
            self.light_sensor.setDeviceSerialNumber(hub_serial_number)
            self.sound_sensor.setDeviceSerialNumber(hub_serial_number)

        # Open channels
        try:
            self.light_sensor.openWaitForAttachment(5000)
            if self.light_sensor:
                logging.info(f"Attached LightSensor successfully!")
                self.light_sensor.setIlluminanceChangeTrigger(0)    # Set illuminance change trigger to 0
                self.light_sensor.setDataInterval(150)           # Data sampling interval, default as 250ms
            
            self.sound_sensor.openWaitForAttachment(5000)
            if self.sound_sensor:
                logging.info(f"Attached SoundSensor successfully!")
                self.sound_sensor.setSPLChangeTrigger(1.0)          # Set SPL change trigger to 1.0
                self.sound_sensor.setDataInterval(150)           # Data sampling interval, default as 250ms
        except PhidgetException as e:
            logging.error("PhidgetException {} ({}): {}".format(e.code, e.codeDescription, e.details))
            return


        self.light_data = []
        self.sound_data = []
        self.timestamps = defaultdict(list)
        self.stop = False

    def init_logging(self):
        """
        Initialize the logging module with the specified configuration.
        """
        logging.basicConfig(level=logging.DEBUG if self.verbose else logging.INFO,
                            format='%(message)s',
                            handlers=[
                                RichHandler(rich_tracebacks=True,                                # Enable rich tracebacks
                                            tracebacks_show_locals=True,                         # Show local variables in tracebacks
                                            log_time_format="[%Y/%m/%d %H:%M:%S]",               # DateTime format
                                            omit_repeated_times=False,                           # Print logging timestamp for each log
                                            keywords=['USB Camera', 'PhidgetIR', 'Hold']),       # Highlight keywords in console
                                logging.FileHandler('Script.log', mode='w', encoding='utf-8')
                            ])

    def onIlluminanceChange(self, handler, illuminance):
        """
        Event handler for illuminance changes in the LightSensor.

        :param handler: The device handler.
        :param illuminance: The illuminance value.
        """
        if illuminance:
            timestamp = datetime.datetime.now()
            self.timestamps['video'].append(timestamp)
            self.light_data.append((timestamp, illuminance))
            logging.debug(f"Timestamp: {timestamp}, Light sensor data: {str(illuminance)}")
        if self.stop:
            self.light_sensor.setOnIlluminanceChangeHandler(None)

    def onSPLChange(self, handler, dB, dBA, dBC, octaves):
        """
        Event handler for SPL changes in the SoundSensor.

        :param handler: The device handler.
        :param dB: The sound level in decibels.
        :param dBA: The A-weighted sound level in decibels.
        :param dBC: The C-weighted sound level in decibels.
        :param octaves: The octave band levels.
        """
        if dB:
            timestamp = datetime.datetime.now()
            self.timestamps['audio'].append(timestamp)
            self.sound_data.append((timestamp, dB))
            # logging.info("{} - Sound sensor data: {}, {}, {}, Octaves: {}".format(
                # timestamp, str(dB), str(dBA), str(dBC), str(octaves[7])))
            logging.debug(f"Timestamp: {timestamp}, Sound sensor data: {str(dB)}")
        if self.stop:
            self.sound_sensor.setOnSPLChangeHandler(None)

    def set_illuminance_data_interval(self, value: int):
        """
        Set the illuminance data interval for the LoundSensor.

        :param interval: The data interval value.
        """
        try:
            if self.light_sensor:
                self.light_sensor.setDataInterval(value)
                logging.info(f"Illuminance Data Interval set to {value}ms")
        except PhidgetException as e:
            logging.error("Failed to set illuminance Data Interval: {}".format(e.details))

    def set_spl_data_interval(self, value: int):
        """
        Set the SPL data interval for the SoundSensor.

        :param interval: The data interval value.
        """
        try:
            if self.sound_sensor:
                self.sound_sensor.setDataInterval(value)
                logging.info(f"SPL Data Interval set to {value}ms")
        except PhidgetException as e:
            logging.error("Failed to set SPL Data Interval: {}".format(e.details))

    def set_illuminance_change_trigger(self, trigger):
        """
        Set the illuminance change trigger for the LightSensor.

        :param illuminanceChangeTrigger: The illuminance change trigger value.
        """
        try:
            if self.light_sensor:
                self.light_sensor.setIlluminanceChangeTrigger(trigger)
                logging.info(f"Illuminance Change Trigger set to {trigger}")
        except PhidgetException as e:
            logging.error("Failed to set Illuminance Change Trigger: {}".format(e.details))

    def set_spl_change_trigger(self, trigger):
        """
        Set the SPL change trigger for the SoundSensor.

        :param splChangeTrigger: The SPL change trigger value.
        """
        try:
            if self.sound_sensor:
                self.sound_sensor.setSPLChangeTrigger(trigger)
                logging.info(f"SPL Change Trigger set to {trigger}")
        except PhidgetException as e:
            logging.error("Failed to set SPL Change Trigger: {}".format(e.details))

    def capture_sensor_data(self, duration):
        """
        Run the test for a specified duration.

        :param duration: The duration of the test in seconds.
        """
        # Assign event handlers
        self.light_sensor.setOnIlluminanceChangeHandler(self.onIlluminanceChange)
        self.sound_sensor.setOnSPLChangeHandler(self.onSPLChange)
        
        start_time = time.time()
        end_time = start_time + duration

        # Use rich.track to visualize the progress
        for _ in track(range(duration), description="Capturing sensor data..."):
            # Check if the loop should stop
            if self.stop or time.time() >= end_time:
                break
            
            # Record timestamps periodically
            timestamp = datetime.datetime.now()
            self.timestamps['video'].append(timestamp)
            self.timestamps['audio'].append(timestamp)
            time.sleep(1)   # Sleep for 1 second to ensure uniform sampling

        self.stop = True    # Ensure the loop exits
        self.close()        # Close sensors to ensure proper exit


    def visualize_data(self, save_picture=None):
        """
        Visualize the collected data from the LightSensor and SoundSensor.
        """
        # Ensure timestamps are not empty
        if not self.timestamps['video']:
            self.timestamps['video'].append(datetime.datetime.now())
        if not self.timestamps['audio']:
            self.timestamps['audio'].append(datetime.datetime.now())

        # Extracting data and aligning based on timestamps
        light_times, light_values = zip(*self.light_data) if self.light_data else ([], [])
        sound_times, sound_values = zip(*self.sound_data) if self.sound_data else ([], [])

        plt.figure(figsize=(15, 10))

        # Plot Light Sensor Data
        ax1 = plt.subplot(2, 1, 1)
        ax1.plot(light_times, light_values, label='Light Sensor')
        ax1.set_xlabel('Time')
        ax1.set_ylabel('Illuminance (lux)')
        ax1.set_title('Light Sensor Data')
        ax1.legend()
        ax1.set_ylim(-10, 400)  # Set Y-axis range for light sensor data, adjust as needed
        ax1.xaxis.set_major_locator(mdates.SecondLocator(interval=1))
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))

        # Calculate and plot median line for Light Sensor Data
        if light_values:
            median_light = np.median(light_values)
            ax1.axhline(y=median_light, color=(1, 0, 0, 0.5), linestyle='--', label='Median')
            ax1.legend()

        # Plot Sound Sensor Data
        ax2 = plt.subplot(2, 1, 2)
        ax2.plot(sound_times, sound_values, label='Sound Sensor')
        ax2.set_xlabel('Time')
        ax2.set_ylabel('Sound Level (dB)')
        ax2.set_title('Sound Sensor Data')
        ax2.legend()
        ax2.set_ylim(0, 120)  # Set Y-axis range for sound sensor data, adjust as needed
        ax2.xaxis.set_major_locator(mdates.SecondLocator(interval=1))
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
        
        # Calculate and plot median line for Sound Sensor Data
        if sound_values:
            median_sound = np.median(sound_values)
            ax2.axhline(y=median_sound, color=(1, 0, 0, 0.5), linestyle='--', label='Median')
            ax2.legend()

        # Rotate date labels for better readability and avoid overlapping
        for ax in [ax1, ax2]:
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')

        plt.tight_layout()
        if save_picture:
            plt.savefig(save_picture)
        plt.show()

    def close(self):
        """
        Close the LightSensor and SoundSensor.
        """
        if self.light_sensor:
            self.light_sensor.close()
        if self.sound_sensor:
            self.sound_sensor.close()

if __name__ == "__main__":
    
    # Usage
    try:
        vinthub = VintHubController(light_sensor_port=1, sound_sensor_port=2, hub_serial_number=751480, verbose=True)
        # controller = VintHubController(light_sensor_port=1, hub_serial_number=751480, verbose=True)
        # controller = VintHubController(sound_sensor_port=2, hub_serial_number=751480, verbose=True)
        vinthub.capture_sensor_data(duration=10)                    # Set test duration as xx seconds
        vinthub.visualize_data(save_picture='sensor_data.png')       # Save the data image as xx.png 
    except PhidgetException as ex:
        traceback.print_exc()
        logging.error("PhidgetException {} ({}): {}".format(ex.code, ex.description, ex.details))