#!/usr/bin/env python
# -*- coding: utf-8 -*-
# -*- Python -*-

"""
 @file DataBaseInPut.py
 @brief DataBaseInPut
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
databaseinput_spec = ["implementation_id", "DataBaseInPut", 
		 "type_name",         "DataBaseInPut", 
		 "description",       "DataBaseInPut", 
		 "version",           "0.0.1", 
		 "vendor",            "Miyamoto Nobuhiko", 
		 "category",          "TEST", 
		 "activity_type",     "STATIC", 
		 "max_instance",      "1", 
		 "language",          "Python", 
		 "lang_type",         "SCRIPT",
                 "conf.default.DataBase_Name", "None",
                 "conf.default.Table_Name", "None",
		 ""]
# </rtc-template>






##
# @class DataBaseInPut
# @brief DataBaseInPut
# 
# 
class DataBaseInPut(OpenRTM_aist.DataFlowComponentBase):
	
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

		self._d_m_in = RTC.TimedStringSeq(RTC.Time(0,0),0)
                self._m_inIn = OpenRTM_aist.InPort("in", self._d_m_in)

                

                                
                self.DataBase_Name = ["None"]
                self.Table_Name = ["None"]
                

                

        		 
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
                
                

                
                self.bindParameter("DataBase_Name", self.DataBase_Name, "None")
                self.bindParameter("Table_Name", self.Table_Name, "None")
                

                
		
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
                        
                  
                        SQL = "INSERT INTO "+ self.Table_Name[0] + " VALUES("
                        for i in range(0,len(data.data)):
                           if i < len(data.data)-1:
                              SQL += "'" + str(data.data[i]) + "',"
                           else:
                              SQL += "'" + str(data.data[i]) + "')"
                        
                        try:
                                print SQL
                                f2 = self._database._ptr().executeUpdate(self.DataBase_Name[0],SQL)
                                
                                if f2 == False:
                                   print "ERROR"
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
	



def DataBaseInPutInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=databaseinput_spec)
    manager.registerFactory(profile,
                            DataBaseInPut,
                            OpenRTM_aist.Delete)

def MyModuleInit(manager):
    DataBaseInPutInit(manager)

    # Create a component
    comp = manager.createComponent("DataBaseInPut")

def main():
	mgr = OpenRTM_aist.Manager.init(sys.argv)
	mgr.setModuleInitProc(MyModuleInit)
	mgr.activateManager()
	mgr.runManager()

if __name__ == "__main__":
	main()

