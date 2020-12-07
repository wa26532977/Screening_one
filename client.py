'''
Open Source Initiative OSI - The MIT License:Licensing
Tue, 2006-10-31 04:56 nelson

The MIT License

Copyright (c) 2009 BK Precision

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

This script talks to the DC load in two ways:
    1.  Using a DCLoad object (you'd use this method when you write a
        python application that talks to the DC load.

    2.  Using the COM interface.  This shows how python code uses the
        COM interface.  Other programming environments (e.g., Visual
        Basic and Visual C++) would use very similar techniques to
        talk to the DC load via COM.

Note that the DCLoad object and the COM server interface functions
always return strings.

$RCSfile: client.py $ 
$Revision: 1.0 $
$Date: 2008/05/16 21:02:50 $
$Author: Don Peterson $
'''

import sys
import dcload
import time

try:
    from win32com.client import Dispatch
except:
    pass


def gettingValue():
    load = Dispatch('BKServers.DCLoad85xx')
    port = "COM3"
    baudrate = "19200"

    load.Initialize(port, baudrate)
    values = load.GetInputValues()
    return values


def turnOff():
    load = Dispatch('BKServers.DCLoad85xx')
    port = "COM3"
    baudrate = "19200"

    def test(cmd, results):
        if results:
            print(cmd + "failed:")
            print(results)
            exit(1)
        else:
            print(cmd)

    load.Initialize(port, baudrate)
    values = load.GetInputValues()
    load.TurnLoadOff()
    test("Set to local control", load.SetLocalControl())
    return values


def GetOCV():
    load = Dispatch('BKServers.DCLoad85xx')
    port = "COM3"
    baudrate = "19200"

    def test(cmd, results):
        if results:
            print(cmd + "failed:")
            print(results)
            exit(1)
        else:
            print(cmd)

    load.Initialize(port, baudrate)
    test("Set to remote control", load.SetRemoteControl())
    test("Set Remote Sense to enable", load.SetRemoteSense(1))
    load.TurnLoadOn()
    values = load.GetInputValues()
    load.TurnLoadOff()
    test("Set to local control", load.SetLocalControl())
    return values


def GetCCV(testing_type, testing_value2, timer):
    cc_value = 0
    load = Dispatch('BKServers.DCLoad85xx')
    port = "COM3"
    baudrate = "19200"

    def test(cmd, results):
        if results:
            print(cmd + "failed:")
            print(results)
            exit(1)
        else:
            print(cmd)

    load.Initialize(port, baudrate)
    test("Set to remote control", load.SetRemoteControl())
    test("Set Remote Sense to enable", load.SetRemoteSense(1))
    load.TurnLoadOn()

    if testing_type == "Constant Current":
        testing_value = testing_value2 / 1000
        test("Set to constant current", load.Setmode("cc"))
        test("Set Transient to CC ", load.SetTransient("cc", 0, 0.01, testing_value, int(timer)*10, "pulse"))
    elif testing_type == "Constant Resistor":
        test("Set to constant current", load.Setmode("cr"))
        test("Set Transient to CC ", load.SetTransient("cr", 0, 0.01, testing_value2, int(timer) * 10, "pulse"))

    test("Set function to Transient", load.SetFunction("transient"))
    load.TurnLoadOn()
    load.TriggerLoad()

    start_time = time.time()
    t_end = time.time() + 4
    values = []
    while time.time() < t_end:
        values.append(load.GetInputValues()[0])
        time.sleep(0.01)
    print("Final values is ")
    print(values)
    print(min(values))
    load.TurnLoadOff()
    test("Set Function to fix", load.SetFunction("fixed"))
    test("Set to local control", load.SetLocalControl())



def TalkToLoad(load, port, baudrate):
    '''load is either a COM object or a DCLoad object.  They have the 
    same interface, so this code works with either.
 
    port is the COM port on your PC that is connected to the DC load.
    baudrate is a supported baud rate of the DC load.
    '''
    def test(cmd, results):
        if results:
            print(cmd + "failed:")
            print(results)
            exit(1)
        else:
            print(cmd)
    # Open a serial connection
    print(load.Initialize(port, baudrate))
    #load.Initialize(port, baudrate)
    test("Set to remote control", load.SetRemoteControl())
    print("Time from DC Load =" + str(load.TimeNow()))
    test("Set to constant current", load.Setmode("cr"))
    #test("Set to remote control", load.SetRemoteControl())
    test("Set max current to 3 A", load.SetMaxCurrent(3))
    test("Set Max Voltage to 15V", load.SetMaxVoltage(15))
    #test("Set CC current to 0.1 A", load.SetCCCurrent(0.03))
    test("Set Transient to CC ", load.SetTransient("cr", 4000, 1, 200, 40, "pulse"))
    test("Set Remote Sense to enable", load.SetRemoteSense(1))
    test("Set function to Transient", load.SetFunction("transient"))
    print("Settings:")
    print("  Transient state     =" + str(load.GetTransient('cr')))
    print("  Mode                =" + str(load.GetMode()))
    print("  Max voltage         =" + str(load.GetMaxVoltage()))
    print("  Max current         =" + str(load.GetMaxCurrent()))
    print("  Max power           =" + str(load.GetMaxPower()))
    print("  CC current          =" + str(load.GetCCCurrent()))
    print("  CV voltage          =" + str(load.GetCVVoltage()))
    print("  CW power            =" + str(load.GetCWPower()))
    print("  CR resistance       =" + str(load.GetCRResistance()))
    print("  Load on timer time  =" + str(load.GetLoadOnTimer()))
    print("  Load on timer state =" + str(load.GetLoadOnTimerState()))
    print("  Trigger source      =" + str(load.GetTriggerSource()))
    print("  Function            =" + str(load.GetFunction()))
    print("  Remote Sense        =" + str(load.GetRemoteSense()))
    print("  Input values:")

    load.TurnLoadOn()
    load.TriggerLoad()
    #print("  TriggerLoad         =" + str(load.TriggerLoad()))


    #test("Set Load on Timer 5 second", load.SetLoadOnTimer(5))

    #values = load.GetInputValues()
    #for i in range(5):
    #    print(values[i])

    #test("Set to constant current", load.Setmode("cc"))

    start_time = time.time()
    t_end = time.time()+11
    values = []
    while time.time() < t_end:
        print(str(time.time()-start_time))
        values.append(load.GetInputValues()[0])
        print(load.GetInputValues())
        time.sleep(0.25)
    print("Final values is ")
    print(values)
    print(min(values))
    load.TurnLoadOff()
    test("Set Function to fix", load.SetFunction("fixed"))
    test("Set to local control", load.SetLocalControl())



def Usage():
    name = sys.argv[0]
    msg = '''Usage:  %(name)s {com|obj} port baudrate
Demonstration python script to talk to a B&K DC load either via the COM
(component object model) interface or via a DCLoad object (in dcload.py).
port is the COM port number on your PC that the load is connected to.  
baudrate is the baud rate setting of the DC load.
''' % locals()
    print(msg)
    exit(1)


if __name__ == '__main__':
    access_type = "com"
    port        = "COM3"
    #baudrate = "9200"
    baudrate    = "19200"
    if access_type == "com":
        load = Dispatch('BKServers.DCLoad85xx')
    elif access_type == "obj":
        load = dcload.DCLoad()
    else:
        Usage()
    TalkToLoad(load, port, baudrate)

