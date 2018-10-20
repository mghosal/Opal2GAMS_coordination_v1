###############################
### RTLAB automation script ###
###############################
### Version: 1.0 ##############

import os
import RtlabApi 
import OpalApiPy
import time
import sys
import glob
import threading
import random
import string
import fncs

realTimeModeList = {'Hardware Synchronized':0, 'Simulation':1, \
                    'Software Synchronized':2, 'Simulation with no data loss':3, \
                    'Simulation with low priority':4}


class RTLAB_automate:
    
    def __init__(self,projectName,modelName):
        self.projectName= projectName
        self.modelName=modelName
        self.instanceID= 0
        self.realTimeMode = realTimeModeList['Software Synchronized']
        self.timeFactor=1.0
        self.acqGroup=1
        
    def Compile(self):
        
        try:
            print "Opening Project"
            projectId = r.OpenProject(self.projectName)
            self.instanceID=projectId
            print "The connection with '%s' is completed." % self.projectName
            
            ## Registering this thread to receive all messages from the controller
              
            scope = 0 # 0 - local settings 1 - global settings
            # Choosing our linux target for compilation
            platform = OpalApiPy.REDHAWK_TARGET
            OpalApiPy.SetTargetPlatform(platform, scope)
            
            # Register to receive compilation messages
            registerType = OpalApiPy.DISPLAY_REGISTER_ALL         
            OpalApiPy.RegisterDisplay(registerType)
              
            ## Start the following compilation steps
            compilationSteps = OpalApiPy.OP_SEPARATE_ONLY | \
                               OpalApiPy.OP_GENERATE_ONLY | \
                               OpalApiPy.OP_TRANSFER_ONLY | \
                               OpalApiPy.OP_COMPIL_ONLY | \
                               OpalApiPy.OP_RETRIEVE_FILES_ONLY | \
                               OpalApiPy.OP_COMPIL_ALL_REDHAWK
             
            print "COMPILING MODEL NOW"
            # Actual compilation call with options to control each subsystems passed as tuples
            OpalApiPy.StartCompile2(( ("",compilationSteps), ))
     
            ## Loop until the end of the compilation
            status = 0       
            while status == 0:
                try:
                    ## Get compilation results
                    res = OpalApiPy.DisplayInformation(1)
                    print res[2]
                    if res[2].find('Compilation duration : ')>0:
                        status = 1                
                    ## wait before 
                    time.sleep(0.1)
                except Exception,exc:
                    ## If a exception occur when calling GetCompileResults function exit
                    info = sys.exc_info()
                    if info[1][0] <> 11:
                        print "An error occured during compilation."
                        status = 2
                        break
                             
             
        except Exception as e:
            print "Exception:",str(e)
            pass
    
    
    def Load(self):
        try:
            print self.projectName
            instanceID=RtlabApi.OpenProject(self.projectName)
            print instanceID
            # Setting the current model in case project has multiple models
            instanceID=RtlabApi.SetCurrentModel(self.modelName)[0]
            print instanceID
            modelState=RtlabApi.GetModelState()[0]
            
            if modelState == RtlabApi.MODEL_LOADABLE:
                RtlabApi.Load(self.realTimeMode, self.timeFactor)
                print "Loading done."
            
            while 1:
                # check if the model is in the paused state after loading
                modelState=RtlabApi.GetModelState()[0]
                if modelState== RtlabApi.MODEL_PAUSED:
                    print "Model has been loaded and paused."
                    break
                        
        except Exception as e:
            print "Error in loading the model:",str(e)
                
    def AcquireToFile(self,enabled, acqGroup, filename="data1", variable="test"):
        try:
            # Get Acquistion control before this can be enabled
            RtlabApi.GetAcquisitionControl(1,acqGroup-1)
            # Enable/Disable the function
            OpalApiPy.SetAcqWriteFile(acqGroup-1, filename, variable, enabled)
            if enabled==0:
                print "Write to file has been disabled."
            else:
                print "Write to file has been disabled."
        except:
            print "Error in setting acquire to file option."
                    
    
    def Acquire(self,acqGroup=1, synchronization = 0,interpolation = 0,threshold = 0,acqTimeStep = 0.1,frameSize = 0):
        
        try:
            # File where the acquisition results are dumped       
            fd=open('output.txt','a')
            stp=1
            #Check if the model is running
            modelState= RtlabApi.GetModelState()[0]
            if (modelState == RtlabApi.MODEL_RUNNING):
                #Reading all values in the acquisition buffer            
                frameSize=0
                while 1:
                    # Get the data for an entire acquisition buffer
                    acqResults = RtlabApi.GetAcqGroupSyncSignals(acqGroup-1, synchronization,interpolation,threshold, acqTimeStep)
                    simSignals,monSignals,simTimeStep,endFrame = acqResults
                    missedData,offset,simTime,sampleSec = monSignals
                    frameSize+=1
                    # Print the last time stamp value in the buffer
                    if endFrame:
                        str="T%d: Last Time Stamp Value in Buffer for Acquisition Group #%d : %f\n" %(acqGroup,acqGroup, simTimeStep)
                        fd.write(str)
                        #print "T%d: Last Time Stamp Value in Buffer for Acquisition Group #%d : %f\n" %(acqGroup,acqGroup, simTimeStep)
                        break
                else :
                    print "Model is in:",modelState
                    
            # Close the file
            fd.close()
        except Exception as e:
            print "Exception:",str(e)
    
    def AcquirePaused(self,signalNames=(),control_values=()):
        try:
            fd1=open('signallist.txt','a')
            modelState=RtlabApi.GetModelState()[0]
            if modelState == RtlabApi.MODEL_PAUSED:
                signalInfos=RtlabApi.GetSignalsDescription()
                for signal in signalInfos:
                    s=str(signal)+"\n"
                    fd1.write(s)
                print "Wrote signal infos to the file"
            else:
                print "Model is running!"
            fd1.close()
        except Exception as e:
            print "Exception:",str(e)
    
    def Snapshot(self,option=0,filename = 'snapshotfile',overwrite = 1,increment = 0,comment = '',commentLen = 0):
        try:
            # First we need to connect to the model
            modelState=RtlabApi.GetModelState()[0]
            
            # We take or retrieve snapshot only when the model is in the paused state
            if modelState == RtlabApi.MODEL_RUNNING:
                
                print "Model is running! Please ensure that the model is paused before calling Snapshot."
            
            elif modelState == RtlabApi.MODEL_PAUSED:
                
                # Get control of the model before you take/retrieve Snapshots
                monitoringControl = 1
                OpalApiPy.GetMonitoringControl(monitoringControl)
                OpalApiPy.Snapshot(option, filename, overwrite, increment, comment, commentLen)
                # Relinquish control
                monitoringControl = 0
                OpalApiPy.GetMonitoringControl(monitoringControl)
                ## 1 = take snapshot, 2 = restore snapshot
                if option ==1:
                    print "Saved a snapshot successfully."
                elif option ==2:
                    print "Retrieved a snapshot successfuly."
                else:
                    print "Invalid Snapshot option. 1 - Take Snapshot, 2 - Retrieve Snapshot"
        except Exception as e:
            monitoringControl = 0
            OpalApiPy.GetMonitoringControl(monitoringControl)
            print "Error in Taking/Retrieving Snapshot:", str(e)
    

if __name__ == "__main__":
                     
    try:
        # This cleans all the model logs, data and other files on the target
        RtlabApi.OP_MODEL_CLEAN_TARGET_RUN                   
        projectName = "C:/Users/Battelle (PNNL)/Desktop/Siddharth/Luigi_Model/Luigi_Model.llp"
        modelName = "C:/Users/Battelle (PNNL)/Desktop/Siddharth/Luigi_Model/models/ssn_IEEE_37bus/ssn_IEEE_37bus.mdl"
        # Create a new object of the RTLAB_automate class
        x=RTLAB_automate(projectName,modelName)
        
        # Compile the model
        
        #x.Compile()
        
        ## Load the current model, optional parameters realtime mode and time factor
        
        x.Load()
        print "Load completed and initializing FNCS..."
        # Initialize FNCS and keep it ready
        config = """name = opalrt
time_delta = 1s
broker = tcp://localhost:5570
values
    bus_701_phA_P
        topic = bus_7XX_phX/bus_701_phA_P
        list=false
    bus_701_phA_P
        topic = bus_7XX_phX/bus_701_phA_P
        list=false
    bus_701_phB_P
        topic = bus_7XX_phX/bus_701_phB_P
        list=false
    bus_701_phC_P
        topic = bus_7XX_phX/bus_701_phC_P 
        list=false
    bus_712_phC_P
        topic = bus_7XX_phX/bus_712_phC_P 
        list=false
    bus_713_phC_P
        topic = bus_7XX_phX/bus_713_phC_P 
        list=false
    bus_714_phA_P
        topic = bus_7XX_phX/bus_714_phA_P
        list=false
    bus_714_phB_P
        topic = bus_7XX_phX/bus_714_phB_P
        list=false
    bus_718_phA_P
        topic = bus_7XX_phX/bus_718_phA_P
        list=false
    bus_720_phC_P
        topic = bus_7XX_phX/bus_720_phC_P
        list=false
    bus_722_phB_P
        topic = bus_7XX_phX/bus_722_phB_P
        list=false
    bus_722_phC_P
        topic = bus_7XX_phX/bus_722_phC_P
        list=false
    bus_724_phB_P
        topic = bus_7XX_phX/bus_724_phB_P
        list=false
    bus_725_phB_P
        topic = bus_7XX_phX/bus_725_phB_P
        list=false
    bus_727_phC_P
        topic = bus_7XX_phX/bus_727_phC_P
        list=false
    bus_701_phA_Q
        topic = bus_7XX_phX/bus_701_phA_Q
        list=false
    bus_701_phB_Q
        topic = bus_7XX_phX/bus_701_phB_Q
        list=false
    bus_701_phC_Q
        topic = bus_7XX_phX/bus_701_phC_Q 
        list=false
    bus_712_phC_Q
        topic = bus_7XX_phX/bus_712_phC_Q 
        list=false
    bus_713_phC_Q
        topic = bus_7XX_phX/bus_713_phC_Q 
        list=false
    bus_714_phA_Q
        topic = bus_7XX_phX/bus_714_phA_Q
        list=false
    bus_714_phB_Q
        topic = bus_7XX_phX/bus_714_phB_Q
        list=false
    bus_718_phA_Q
        topic = bus_7XX_phX/bus_718_phA_Q
        list=false
    bus_720_phC_Q
        topic = bus_7XX_phX/bus_720_phC_Q
        list=false
    bus_722_phB_Q
        topic = bus_7XX_phX/bus_722_phB_Q
        list=false
    bus_722_phC_Q
        topic = bus_7XX_phX/bus_722_phC_Q
        list=false
    bus_724_phB_Q
        topic = bus_7XX_phX/bus_724_phB_Q
        list=false
    bus_725_phB_Q
        topic = bus_7XX_phX/bus_725_phB_Q
        list=false
    bus_727_phC_Q
        topic = bus_7XX_phX/bus_727_phC_Q
        list=false"""
        
        voltages_topic = ["bus_701_phA_V","bus_701_phB_V","bus_701_phC_V","bus_712_phC_V","bus_713_phC_V","bus_714_phA_V","bus_714_phB_V","bus_718_phA_V","bus_720_phC_V","bus_722_phB_V","bus_722_phC_V","bus_724_phB_V","bus_725_phB_V","bus_727_phC_V"]
        load_p_topic = ["bus_701_phA_P","bus_701_phB_P","bus_701_phC_P","bus_712_phC_P","bus_713_phC_P","bus_714_phA_P","bus_714_phB_P","bus_718_phA_P","bus_720_phC_P","bus_722_phB_P","bus_722_phC_P","bus_724_phB_P","bus_725_phB_P","bus_727_phC_P"]
        load_q_topic = ["bus_701_phA_Q","bus_701_phB_Q","bus_701_phC_Q","bus_712_phC_Q","bus_713_phC_Q","bus_714_phA_Q","bus_714_phB_Q","bus_718_phA_Q","bus_720_phC_Q","bus_722_phB_Q","bus_722_phC_Q","bus_724_phB_Q","bus_725_phB_Q","bus_727_phC_Q"]
        parameterNames_P=('ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load1/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load2/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L712/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L713/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load1/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L718/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L720/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load1/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L724/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L725/Single-Phase Static Load13/GLD_pNom/Value','ssn_IEEE_37bus/sm_computation/L727/Single-Phase Static Load13/GLD_pNom/Value')
        parameterNames_Q=('ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load1/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load2/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L712/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L713/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load1/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L718/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L720/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load1/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L724/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L725/Single-Phase Static Load13/GLD_qNom/Value','ssn_IEEE_37bus/sm_computation/L727/Single-Phase Static Load13/GLD_qNom/Value')
        signalNames=('ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load1/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load1/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load2/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L701/Single-Phase Static Load2/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L712/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L712/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L713/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L713/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load1/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L714/Single-Phase Static Load1/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L718/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L718/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L720/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L720/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load1/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L722/Single-Phase Static Load1/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L724/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L724/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L725/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L725/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)','ssn_IEEE_37bus/sm_computation/L727/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(1)','ssn_IEEE_37bus/sm_computation/L727/Single-Phase Static Load13/Magnitude-Angle to Complex/port1(2)')
        #Initializing FNCS
        fncs.initialize(config)
        print "FNCS initialized.....\n" + str(RtlabApi.GetModelState()[0])
        # This the logic that controls the execution, acquisition and controls
        realTimeMode=realTimeModeList['Hardware Synchronized']
        timeFactor=1.0
        PauseTime=1.0
        multiplier=1
        if realTimeMode==1:
            PauseTime= PauseTime*multiplier
        # We are going to pause and resume for five different steps and at each time we need to acquire, and send control back to the model.
        loop = 0
        RtlabApi.GetSystemControl()
        RtlabApi.SetPauseTime(PauseTime-0.00)
        RtlabApi.SetStopTime(25*PauseTime-0.00)
        (modelState, realTimeMode) = RtlabApi.GetModelState()
        if modelState == RtlabApi.MODEL_PAUSED:
            start_time=time.time()
            # Execute the model until the next pause time
            print "Starting Execution"
        print "Starting Execution...\n"
        timecontrol=0
        while loop < 2:
            loop+=1
            RtlabApi.Execute(1)
            flag=0
            # Keep running the simulation and acquire the data from the OPAL target asynchronously        
            while timecontrol<5:
                timecontrol=timecontrol+1
                current_time=time.time()
                elapsed_time=(current_time-start_time)
                print "Time elapsed at Pause:",elapsed_time
                # Request time from FNCS broker to publish data from OPAL Target into FNCS
                print int(elapsed_time)
                fncs_current_time = fncs.time_request(int(elapsed_time)+1)
                # Read and publish all voltages
                #signals_l is list of all signals with different parameters
                signals_l=[]
                signalValues=()
                signals_l = list(RtlabApi.GetSignalsDescription())
                #find only V signal values of interest
                for sig_name in signalNames:
                    for item in signals_l:
                        if sig_name in item:
                            signalValues=signalValues+(item[6],)
                i=0
                for voltage in voltages_topic:
                    Voltage_real=signalValues[2*i]
                    Voltage_imag=signalValues[2*i+1]
                    if Voltage_imag<0:
                        Voltage_str=str(Voltage_real) + str(Voltage_imag) + "j"
                    else:
                        Voltage_str=str(Voltage_real) + "+" + str(Voltage_imag) + "j"
                    fncs.publish(voltage,Voltage_str)
                    i+=1
                #Get P&Q values
                P_val=()
                for P in load_p_topic:
                    P_val=P_val+(round((float(fncs.get_value(P))/10000),2),)
                
                print P_val
                Q_val=()
                for Q in load_q_topic:
                    Q_val=Q_val+(round((float(fncs.get_value(Q))/10000),2),)
                print Q_val
                # Get control to send parameters
                RtlabApi.GetParameterControl(1)
                # Send control signals after getting control
                timetest=time.time()
                RtlabApi.SetParametersByName(parameterNames_Q+parameterNames_P,Q_val+P_val)
                print time.time()-timetest
                RtlabApi.GetParameterControl(0)
                if (time.time()-current_time<1):
                    print "Sleep time was", time.time()-current_time
                    time.sleep(1-(time.time()-current_time))
                    
                #End of while loop
        end_time=time.time()
        elapsed_time=end_time-start_time
        print "Total Elapsed time:",elapsed_time,"sec...."
        #fncs.finalize()
        print "Done!"

    except Exception as e:
        print "Problem:",e
        RtlabApi.Reset()
        print "The model is reset"
        RtlabApi.Disconnect()
        print "The connection is closed."

    finally:
        RtlabApi.Reset()
        print "The model is reset"
        RtlabApi.Disconnect()
        print "The connection is closed."
