# MiscTools
Includes various test scripts

# JiraViz
Collect Jira related script and get more help in daily work. 
进入学习模式1(阻塞模式, 依赖参数:[-l, -c CUSTOMER]): 接收红外键值. ex: python3 iRemote.py -l -c Hisense
ir.learn_code()
    
进入学习模式2(阻塞模式, 依赖参数:[-m, -c CUSTOMER]): 接收红外键值. ex: python3 iRemote.py -m -c Hisense
ir.learn_code2()

进入单个键值发射模式(单次执行, 依赖参数:[-s, -c CUSTOMER, -k KEY]): 发送红外键值. ex: python3 iRemote.py -s -c Hisense -k Home
ir.transmit_code(*ir.code_transition(config.customer, config.key))

通过[-r, -c]组合来控制随机发送指定厂商遥控键值
random_transmit(manufacturer=config.customer)
        
# PhidgetIR
Universal infrared remote control.
if LightSensor x1 and SoundSensor x1 exist
vinthub = VintHubController(light_sensor_port=1, sound_sensor_port=2, hub_serial_number=751480)
vinthub.capture_sensor_data(duration=10, light_threshold=10, sound_threshold=50)                    # Set test duration as xx seconds, light and sound warning threshold 
vinthub.visualize_data(save_picture='sensor_data.png')                                              # Save the data image as xx.png 
        
if only LightSensor x1 exist
vinthub = VintHubController(light_sensor_port=1, hub_serial_number=751480)
vinthub.capture_sensor_data(duration=10, light_threshold=10)                                      # Set test duration as xx seconds, light warning threshold 
vinthub.visualize_data(save_picture='sensor_data.png')                                            # Save the data image as xx.png 
        
if only SoundSensor x1 exist
vinthub = VintHubController(sound_sensor_port=2, hub_serial_number=751480)
vinthub.capture_sensor_data(duration=10, sound_threshold=10)                                      # Set test duration as xx seconds, sound warning threshold 
vinthub.visualize_data(save_picture='sensor_data.png')                                            # Save the data image as xx.png 
