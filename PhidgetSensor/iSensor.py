# !/usr/bin/python
# -*- coding: UTF-8 -*-
##################################################
# __Date__: 2024/5/30
# __Python__: 3.9.6
# __Author__: Jianfan.Ai
##################################################

import os
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

def init_logging(verbose):
    """
    Initialize the logging module with the specified configuration.
    """
    # Create a logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Create handlers
    console_handler = RichHandler(rich_tracebacks=True,                               # Enable rich tracebacks
                                  tracebacks_show_locals=True,                        # Show local variables in tracebacks
                                  log_time_format="[%Y/%m/%d %H:%M:%S]",              # DateTime format
                                  omit_repeated_times=False,                          # Print logging timestamp for each log
                                  keywords=[""])                                      # Highlight keywords in console
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)              # Set level for console handler

    file_handler = logging.FileHandler('Script.log', mode='a', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)  # Set level for file handler

    # Create formatters
    console_formatter = logging.Formatter('%(message)s')
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y/%m/%d %H:%M:%S')

    # Add formatters to handlers
    console_handler.setFormatter(console_formatter)
    file_handler.setFormatter(file_formatter)

    # Add handlers to the logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

class VintHubController:
    def __init__(self, light_sensor_port: int, sound_sensor_port: int, hub_serial_number: int):
        """
        Initialize the VintHubController with specified ports and serial number.

        :param light_sensor_port: The hub port for the LightSensor.
        :param sound_sensor_port: The hub port for the SoundSensor.
        :param hub_serial_number: The serial number for the VINT Hub.
        """
        # Initialize sensors as None
        self.light_sensor = None
        self.sound_sensor = None

        # Initialize LightSensor if port is provided
        if light_sensor_port is not None:
            self.light_sensor = LightSensor()
            self.light_sensor.setHubPort(light_sensor_port)  # LightSensor Port
            if hub_serial_number:
                self.light_sensor.setDeviceSerialNumber(hub_serial_number)
            try:
                self.light_sensor.openWaitForAttachment(5000)
                logging.info("Attached LightSensor successfully!")
                self.light_sensor.setIlluminanceChangeTrigger(0)  # Set illuminance change trigger to 0
                self.light_sensor.setDataInterval(150)  # Data sampling interval, default as 250ms
            except PhidgetException as e:
                logging.error("PhidgetException {} ({}): {}".format(e.code, e.codeDescription, e.details))

        # Initialize SoundSensor if port is provided
        if sound_sensor_port is not None:
            self.sound_sensor = SoundSensor()
            self.sound_sensor.setHubPort(sound_sensor_port)  # SoundSensor Port
            if hub_serial_number:
                self.sound_sensor.setDeviceSerialNumber(hub_serial_number)
            try:
                self.sound_sensor.openWaitForAttachment(5000)
                logging.info("Attached SoundSensor successfully!")
                self.sound_sensor.setSPLChangeTrigger(1.0)  # Set SPL change trigger to 1.0
                self.sound_sensor.setDataInterval(150)  # Data sampling interval, default as 250ms
            except PhidgetException as e:
                logging.error("PhidgetException {} ({}): {}".format(e.code, e.codeDescription, e.details))

        self.light_data = []
        self.sound_data = []
        self.timestamps = defaultdict(list)
        self.stop = False

    def onIlluminanceChange(self, handler, illuminance):
        """
        Event handler for illuminance changes in the LightSensor.

        :param handler: The device handler.
        :param illuminance: The illuminance value.
        """
        if illuminance > 0:
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
        if dB > 0:
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

    def capture_sensor_data(self, duration, light_threshold=None, sound_threshold=None):
        """
        Run the test for a specified duration and check for anomalies in sensor data.

        :param duration: The duration of the test in seconds.
        :param light_threshold: The threshold value for the light sensor.
        :param sound_threshold: The threshold value for the sound sensor.
        :return: Dictionary with anomaly detection results.
        """
        # Assign event handlers
        if self.light_sensor:
            self.light_sensor.setOnIlluminanceChangeHandler(self.onIlluminanceChange)
        if self.sound_sensor:
            self.sound_sensor.setOnSPLChangeHandler(self.onSPLChange)
        
        start_time = time.time()
        end_time = start_time + duration

        anomalies = {'light': False, 'sound': False}

        # Use rich.track to visualize the progress
        for _ in track(range(duration), description="Capturing sensor data..."):
            # Check if the loop should stop
            if self.stop or time.time() >= end_time:
                break
            
            # Record timestamps periodically
            timestamp = datetime.datetime.now()

            if self.light_data:
                self.timestamps['video'].append(timestamp)
            if self.sound_data:
                self.timestamps['audio'].append(timestamp)
            
            # Check for anomalies
            if light_threshold:
                if self.light_data and self.light_data[-1][1] < light_threshold:
                    anomalies['light'] = True
                    logging.warning(f"Anomaly detected in light sensor data at {timestamp}: {self.light_data[-1][1]} < {light_threshold}")
            
            if sound_threshold:
                if self.sound_data and self.sound_data[-1][1] < sound_threshold:
                    anomalies['sound'] = True
                    logging.warning(f"Anomaly detected in sound sensor data at {timestamp}: {self.sound_data[-1][1]} < {sound_threshold}")
            
            time.sleep(1)   # Sleep for 1 second to ensure uniform sampling

        self.stop = True    # Ensure the loop exits
        self.close()        # Close sensors to ensure proper exit

        return anomalies


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

        if self.light_data:
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
            
            # Rotate date labels for better readability and avoid overlapping
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        if self.sound_data:
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
            plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45, ha='right')

        plt.tight_layout()
        if save_picture:
            plt.savefig(save_picture)
            logging.info(f"Save captured picture as: {os.path.join(os.getcwd(), save_picture)}")
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
    init_logging(verbose=True)
    
    try:
        # if LightSensor x1 and SoundSensor x1 exist
        vinthub = VintHubController(light_sensor_port=1, sound_sensor_port=2, hub_serial_number=751480)
        vinthub.capture_sensor_data(duration=10, light_threshold=10, sound_threshold=50)                    # Set test duration as xx seconds, light and sound warning threshold 
        vinthub.visualize_data(save_picture='sensor_data.png')                                              # Save the data image as xx.png 
        
        # if only LightSensor x1 exist
        # vinthub = VintHubController(light_sensor_port=1, hub_serial_number=751480)
        # vinthub.capture_sensor_data(duration=10, light_threshold=10)                                      # Set test duration as xx seconds, light warning threshold 
        # vinthub.visualize_data(save_picture='sensor_data.png')                                            # Save the data image as xx.png 
        
        # if only SoundSensor x1 exist
        # vinthub = VintHubController(sound_sensor_port=2, hub_serial_number=751480)
        # vinthub.capture_sensor_data(duration=10, sound_threshold=10)                                      # Set test duration as xx seconds, sound warning threshold 
        # vinthub.visualize_data(save_picture='sensor_data.png')                                            # Save the data image as xx.png 
        
    except PhidgetException as ex:
        traceback.print_exc()
        logging.error("PhidgetException {} ({}): {}".format(ex.code, ex.description, ex.details))