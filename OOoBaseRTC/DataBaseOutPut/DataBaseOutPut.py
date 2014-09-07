#!/usr/bin/env python
# -*- coding: utf-8 -*-
# -*- Python -*-

"""
 @file DataBaseOutPut.py
 @brief DataBaseOutPut
 @date $Date$


"""
import sys
import time
sys.path.append(".")

# Import RTM module
import RTC
import OpenRTM_aist

import DataBase_idl

import random

# Import Service implementation class
# <rtc-template block="service_impl">

# </rtc-template>

# Import Service stub modules
# <rtc-template block="consumer_import">
import DataBase, DataBase__POA


# </rtc-template>


# This module's spesification
# <rtc-template block="module_spec">
databaseoutput_spec = ["implementation_id", "DataBaseOutPut", 
		 "type_name",         "DataBaseOutPut", 
		 "description",       "DataBaseOutPut", 
		 "version",           "0.0.1", 
		 "vendor",            "Miyamoto Nobuhiko", 
		 "category",          "TEST", 
		 "activity_type",     "STATIC", 
		 "max_instance",      "1", 
		 "language",          "Python", 
		 "lang_type",         "SCRIPT",
                 "conf.default.DataBase_Name", "None",
                 "conf.default.Table_Name", "None",
                 "conf.default.row_out", "None",
                 "conf.default.row_in", "None",
                 "conf.default.OutPortDataType", "TimedStringSeq",
                 "conf.default.InPortDataType", "TimedString",
                 "conf.__widget__.DataBase_Name", "text",
                 "conf.__widget__.Table_Name", "text",
                 "conf.__widget__.row_out", "text",
                 "conf.__widget__.row_in", "text",
                 "conf.__widget__.OutPortDataType", "radio",
                 "conf.__widget__.InPortDataType", "radio",
                 "conf.__constraints__.OutPortDataType", """(TimedDoubleSeq,TimedLongSeq,TimedFloatSeq,TimedShortSeq,TimedULongSeq,TimedUShortSeq,TimedCharSeq,TimedWCharSeq,
                                                            TimedBooleanSeq,TimedOctetSeq,TimedStringSeq,TimedWStringSeq)""",
                 "conf.__constraints__.InPortDataType", """(TimedDouble,TimedLong,TimedFloat,TimedShort,TimedULong,TimedUShort,TimedChar,TimedWChar,
                                                            TimedBoolean,TimedOctet,TimedString,TimedWString)""",
		 ""]
# </rtc-template>


##
# コンフィギュレーションパラメータが更新されたときのコールバック
##

class MyConfigUpdateParam(OpenRTM_aist.ConfigurationSetListener):
   def __init__(self,e_rtc):
        self.m_rtc =  e_rtc
   def __call__(self, config_param_name):
        self.m_rtc.ConfigUpdate()



##
# @class DataBaseOutPut
# @brief DataBaseOutPut
# 
# 
class DataBaseOutPut(OpenRTM_aist.DataFlowComponentBase):
	
	##
	# @brief constructor
	# @param manager Maneger Object
	# 
	def __init__(self, manager):
		OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)


		"""
		"""
		self._DataBasePort = OpenRTM_aist.CorbaPort("DataBase")

		

		"""
		"""
		self._database = OpenRTM_aist.CorbaConsumer(interfaceType=DataBase.mDataBase)

		# initialize of configuration-data.
		# <rtc-template block="init_conf_param">
		
		# </rtc-template>

		self._d_m_in = RTC.TimedString(RTC.Time(0,0),0)
                self._m_inIn = OpenRTM_aist.InPort("in", self._d_m_in)

                self._d_m_out = RTC.TimedStringSeq(RTC.Time(0,0),0)
                self._m_outOut = OpenRTM_aist.OutPort("out", self._d_m_out)

                                
                self.DataBase_Name = ["None"]
                self.Table_Name = ["None"]
                self.row_out = ["None"]
                self.row_in = ["None"]
                self.OutPortDataType = ["TimedStringSeq"]
                self.InPortDataType = ["TimedString"]

                self.currentOutPortDataType = "TimedStringSeq"
                self.currentInPortDataType = "TimedString"

                self.DSName = "None"

        ##
        # コンフィギュレーションパラメータが変更されたときに呼び出される関数
        ##
  
        def ConfigUpdate(self):
                
                self._configsets.update("default")
                
                
                if self.currentInPortDataType != self.InPortDataType[0]:
                        self._m_inIn.disconnect_all()
                        self.removePort(self._m_inIn)

                        if self.InPortDataType[0] == "TimedDouble":
                                self._d_m_in = RTC.TimedDouble(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedLong":
                                self._d_m_in = RTC.TimedLong(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedFloat":
                                self._d_m_in = RTC.TimedFloat(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedShort":
                                self._d_m_in = RTC.TimedShort(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedULong":
                                self._d_m_in = RTC.TimedULong(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedChar":
                                self._d_m_in = RTC.TimedChar(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedWChar":
                                self._d_m_in = RTC.TimedWChar(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedBoolean":
                                self._d_m_in = RTC.TimedBoolean(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedOctet":
                                self._d_m_in = RTC.TimedOctet(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedString":
                                self._d_m_in = RTC.TimedString(RTC.Time(0,0),0)
                        elif self.InPortDataType[0] == "TimedWString":
                                self._d_m_in = RTC.TimedWString(RTC.Time(0,0),0)
                                                
                        self._m_inIn = OpenRTM_aist.InPort("in", self._d_m_in)
                        self.addInPort("in",self._m_inIn)
                        self.currentInPortDataType = self.InPortDataType[0]
                        
                if self.currentOutPortDataType != self.OutPortDataType[0]:
                        self._m_outOut.disconnect_all()
                        self.removePort(self._m_outOut)

                        if self.OutPortDataType[0] == "TimedDoubleSeq":
                                self._d_m_out = RTC.TimedDoubleSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedLongSeq":
                                self._d_m_out = RTC.TimedLongSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedFloatSeq":
                                self._d_m_out = RTC.TimedFloatSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedShortSeq":
                                self._d_m_out = RTC.TimedShortSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedULongSeq":
                                self._d_m_out = RTC.TimedULongSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedCharSeq":
                                self._d_m_out = RTC.TimedCharSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedWCharSeq":
                                self._d_m_out = RTC.TimedWCharSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedBooleanSeq":
                                self._d_m_out = RTC.TimedBooleanSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedOctetSeq":
                                self._d_m_out = RTC.TimedOctetSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedStringSeq":
                                self._d_m_out = RTC.TimedStringSeq(RTC.Time(0,0),0)
                        elif self.OutPortDataType[0] == "TimedWStringSeq":
                                self._d_m_out = RTC.TimedWStringSeq(RTC.Time(0,0),0)

                        
                        self._m_outOut = OpenRTM_aist.OutPort("out", self._d_m_out)
                        self.addOutPort("out",self._m_outOut)
                        self.currentOutPortDataType = self.OutPortDataType[0]
                

		 
	##
	#
	# The initialize action (on CREATED->ALIVE transition)
	# formaer rtc_init_entry() 
	# 
	# @return RTC::ReturnCode_t
	# 
	#
	def onInitialize(self):
		# Bind variables and configuration variable
		
		# Set InPort buffers
		
		# Set OutPort buffers
		
		# Set service provider to Ports
		
		# Set service consumers to Ports
		self._DataBasePort.registerConsumer("database", "DataBase::mDataBase", self._database)
		
		# Set CORBA Service Ports
		self.addPort(self._DataBasePort)

		self.addInPort("in",self._m_inIn)
                self.addOutPort("out",self._m_outOut)
                

                self.addConfigurationSetListener(OpenRTM_aist.ConfigurationSetListenerType.ON_SET_CONFIG_SET, MyConfigUpdateParam(self))

                self.bindParameter("DataBase_Name", self.DataBase_Name, "None")
                self.bindParameter("Table_Name", self.Table_Name, "None")
                self.bindParameter("row_out", self.row_out, "None")
                self.bindParameter("row_in", self.row_in, "None")
                self.bindParameter("OutPortDataType", self.OutPortDataType, "TimedStringSeq")
                self.bindParameter("InPortDataType", self.InPortDataType, "TimedString")

                
		
		return RTC.RTC_OK
	
	#	##
	#	# 
	#	# The finalize action (on ALIVE->END transition)
	#	# formaer rtc_exiting_entry()
	#	# 
	#	# @return RTC::ReturnCode_t
	#
	#	# 
	#def onFinalize(self, ec_id):
	#
	#	return RTC.RTC_OK
	
	#	##
	#	#
	#	# The startup action when ExecutionContext startup
	#	# former rtc_starting_entry()
	#	# 
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onStartup(self, ec_id):
	#
	#	return RTC.RTC_OK
	
	#	##
	#	#
	#	# The shutdown action when ExecutionContext stop
	#	# former rtc_stopping_entry()
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onShutdown(self, ec_id):
	#
	#	return RTC.RTC_OK
	
		##
		#
		# The activated action (Active state entry action)
		# former rtc_active_entry()
		#
		# @param ec_id target ExecutionContext Id
		# 
		# @return RTC::ReturnCode_t
		#
		#
	def onActivated(self, ec_id):
                flag = True
		while flag:
                        try:
                                if self._database._ptr().setConnection(self.DataBase_Name[0],"",""):
                                        flag = False       

                        except:
                                pass
                
		return RTC.RTC_OK
	
	#	##
	#	#
	#	# The deactivated action (Active state exit action)
	#	# former rtc_active_exit()
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onDeactivated(self, ec_id):
	#
	#	return RTC.RTC_OK
	
		##
		#
		# The execution action that is invoked periodically
		# former rtc_active_do()
		#
		# @param ec_id target ExecutionContext Id
		#
		# @return RTC::ReturnCode_t
		#
		#
	def onExecute(self, ec_id):
                if self._m_inIn.isNew():
                        data = self._m_inIn.read()
                        self.DSName = self.DataBase_Name[0] + self.Table_Name[0] + str(data.tm.sec) + str(data.tm.nsec)
                        SQL = "SELECT "+ self.row_out[0] + " FROM " + self.Table_Name[0] + " WHERE " + self.row_in[0] + " = " + "'" + str(data.data) + "'"
                        
                        try:
                                f2 = self._database._ptr().executeQuery(self.DSName, self.DataBase_Name[0],SQL)
                                print SQL
                                if f2:
                                        
                                        self._d_m_out.data = []
                                        flag = True
                                        
                                        while flag:
                                                if self._database._ptr().ResultSetNext(self.DSName):
                                                        if self.OutPortDataType[0] == "TimedDoubleSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getDouble(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedLongSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getLong(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedFloatSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getFloat(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedShortSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getShort(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedULongSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getLong(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedCharSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getByte(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedWCharSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getByte(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedBooleanSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getBoolean(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedOctetSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getByte(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedStringSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getString(self.DSName,1))
                                                        elif self.OutPortDataType[0] == "TimedWStringSeq":
                                                                self._d_m_out.data.append(self._database._ptr().getString(self.DSName,1))

                                                else:
                                                        flag = False

                                        
                                        OpenRTM_aist.setTimestamp(self._d_m_out)
                                        self._m_outOut.write()
                        except:
                                print "ERROR"
                
		return RTC.RTC_OK
	
	#	##
	#	#
	#	# The aborting action when main logic error occurred.
	#	# former rtc_aborting_entry()
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onAborting(self, ec_id):
	#
	#	return RTC.RTC_OK
	
	#	##
	#	#
	#	# The error action in ERROR state
	#	# former rtc_error_do()
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onError(self, ec_id):
	#
	#	return RTC.RTC_OK
	
	#	##
	#	#
	#	# The reset action that is invoked resetting
	#	# This is same but different the former rtc_init_entry()
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onReset(self, ec_id):
	#
	#	return RTC.RTC_OK
	
	#	##
	#	#
	#	# The state update action that is invoked after onExecute() action
	#	# no corresponding operation exists in OpenRTm-aist-0.2.0
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#

	#	#
	#def onStateUpdate(self, ec_id):
	#
	#	return RTC.RTC_OK
	
	#	##
	#	#
	#	# The action that is invoked when execution context's rate is changed
	#	# no corresponding operation exists in OpenRTm-aist-0.2.0
	#	#
	#	# @param ec_id target ExecutionContext Id
	#	#
	#	# @return RTC::ReturnCode_t
	#	#
	#	#
	#def onRateChanged(self, ec_id):
	#
	#	return RTC.RTC_OK
	



def DataBaseOutPutInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=databaseoutput_spec)
    manager.registerFactory(profile,
                            DataBaseOutPut,
                            OpenRTM_aist.Delete)

def MyModuleInit(manager):
    DataBaseOutPutInit(manager)

    # Create a component
    comp = manager.createComponent("DataBaseOutPut")

def main():
	mgr = OpenRTM_aist.Manager.init(sys.argv)
	mgr.setModuleInitProc(MyModuleInit)
	mgr.activateManager()
	mgr.runManager()

if __name__ == "__main__":
	main()

