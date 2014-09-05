# -*- coding: utf-8 -*-

import optparse
import sys,os,platform
import codecs
import re
import time
import random
import commands
import math



if os.name == 'posix':
    sys.path += ['./OOoRTC', './OOoRTC/BaseIDL', '/usr/lib/python2.6/dist-packages', '/usr/lib/python2.6/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\BaseIDL', 'C:\\Python26\\lib\\site-packages', 'C:\\Python26\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python26\\lib\\site-packages\\rtctree\\rtmidl']
    
    

import time
import random
import commands
import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming





import uno
import unohelper
import traceback
from com.sun.star.awt import Rectangle
from com.sun.star.beans import PropertyValue

from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext


from  com.sun.star.sdbc.DataType import BIT
from  com.sun.star.sdbc.DataType import TINYINT
from  com.sun.star.sdbc.DataType import SMALLINT
from  com.sun.star.sdbc.DataType import INTEGER
from  com.sun.star.sdbc.DataType import BIGINT
from  com.sun.star.sdbc.DataType import FLOAT
from  com.sun.star.sdbc.DataType import REAL
from  com.sun.star.sdbc.DataType import DOUBLE
from  com.sun.star.sdbc.DataType import NUMERIC
from  com.sun.star.sdbc.DataType import DECIMAL
from  com.sun.star.sdbc.DataType import CHAR
from  com.sun.star.sdbc.DataType import VARCHAR
from  com.sun.star.sdbc.DataType import LONGVARCHAR
from  com.sun.star.sdbc.DataType import DATE
from  com.sun.star.sdbc.DataType import TIME
from  com.sun.star.sdbc.DataType import TIMESTAMP
from  com.sun.star.sdbc.DataType import BINARY
from  com.sun.star.sdbc.DataType import VARBINARY
from  com.sun.star.sdbc.DataType import LONGVARBINARY
from  com.sun.star.sdbc.DataType import SQLNULL
from  com.sun.star.sdbc.DataType import OTHER




import OOoRTC

import DataBase_idl
from omniORB import PortableServer
import DataBase, DataBase__POA
import _GlobalIDL, _GlobalIDL__POA




#comp_num = random.randint(1,3000)
imp_id = "OOoBaseControl"# + str(comp_num)







ooobasecontrol_spec = ["implementation_id", imp_id,
                  "type_name",         imp_id,
                  "description",       "Openoffice Base Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  ""]

##
#サービスポートDataBase
##
class mDataBase_i (DataBase__POA.mDataBase):
    """
    @class mDataBase_i
    Example class implementing IDL interface DataBase.mDataBase
    """

    def __init__(self):
        """
        @brief standard constructor
        Initialise member variables here
        """
        
        
        

    # void setConnection(in string name, in string usr_name, in string passward)
    def setConnection(self, name, usr_name, passward):
        if OOoRTC.base_comp.ConnectionList.has_key(name):
          return True
        try:
          tmp = {}
          db = OOoRTC.base_comp.base._context.getByName(name)
          tmp["Connection"] = db.getConnection(usr_name,passward)
          tmp["Statement"] = tmp["Connection"].createStatement()
          
          OOoRTC.base_comp.ConnectionList[name] = tmp
          return True
        except:
          return False
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None

    # void executeQuery(in string fn)
    def executeQuery(self, name, con, oSQL):
        if OOoRTC.base_comp.ConnectionList.has_key(con):
          try:
            OOoRTC.base_comp.ResultSet[name] = OOoRTC.base_comp.ConnectionList[con]["Statement"].executeQuery(oSQL)
            return True
          except:
            return False
          raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        else:
          return False
        # *** Implement me
        # Must return: None

    # boolean ResultSetNext()
    def ResultSetNext(self, name):
        
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].next()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean ResultSetPrevious()
    def ResultSetPrevious(self, name):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].previous()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean ResultSetFirst()
    def ResultSetFirst(self, name):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].first()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean ResultSetLast()
    def ResultSetLast(self, name):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].last()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean ResultSetBeforeFirst()
    def ResultSetBeforeFirst(self, name):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].beforeFirst()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean ResultSetAfterLast()
    def ResultSetAfterLast(self, name):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].afterLast()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # octet getByte(in short num)
    def getByte(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return OOoRTC.base_comp.ResultSet[name].getByte(num)
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # short getShort(in short num)
    def getShort(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return int(OOoRTC.base_comp.ResultSet[name].getShort(num))
        return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # long getLong(in short num)
    def getLong(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return long(OOoRTC.base_comp.ResultSet[name].getLong(num))
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # float getFloat(in short num)
    def getFloat(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return float(OOoRTC.base_comp.ResultSet[name].getFloat(num))
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # double getDouble(in short num)
    def getDouble(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return float(OOoRTC.base_comp.ResultSet[name].getDouble(num))
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean getBoolean(in short num)
    def getBoolean(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return bool(OOoRTC.base_comp.ResultSet[name].getBoolean(num))
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # string getString(in short num)
    def getString(self, name, num):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return str(OOoRTC.base_comp.ResultSet[name].getString(num))
        else:
          return ""
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # StringLine getDataBaseNames()
    def getDataBaseNames(self):
        Ans = []
        names = OOoRTC.base_comp.base._context.getElementNames()
        for i in names:
            Ans.append(str(i))
        return Ans
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # StringLine getDataTableNames(in string con)
    def getDataTableNames(self, con):
        Ans = []
        
        
        
        if OOoRTC.base_comp.ConnectionList.has_key(con):
            try:
		
                oDBTables = OOoRTC.base_comp.ConnectionList[con]["Connection"].getTables().createEnumeration()
                while oDBTables.hasMoreElements():
                    oTable = oDBTables.nextElement()
                    Ans.append(str(oTable.Name))
                    
            except:
                Ans.append("ERROR")
                

        return Ans
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean executeUpdate(in string name, in string oSQL)
    def executeUpdate(self, con, oSQL):
        
        oRstDataSources = OOoRTC.base_comp.base._context.getByName(con)
        
        
        if OOoRTC.base_comp.ConnectionList.has_key(con):
          try:

            OOoRTC.base_comp.ConnectionList[con]["Statement"].executeUpdate(oSQL)
            
            return True
          except:
            return False
          raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        else:
          return False
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # short getRow(in string name)
    def getRow(self, name):
        if OOoRTC.base_comp.ResultSet.has_key(name):
          return int(OOoRTC.base_comp.ResultSet[name].getRow())
        return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    def AddTable(self, name, con, cols, dt):
        if OOoRTC.base_comp.ConnectionList.has_key(con):
          try:
            oTables = OOoRTC.base_comp.ConnectionList[con]["Connection"].getTables()
            oDBTables = oTables.createEnumeration()

            while oDBTables.hasMoreElements():
                oTable = oDBTables.nextElement()
                if str(oTable.Name) == name:
                    return False

            

            oTableDescriptor = oTables.createDataDescriptor() 
            oTableDescriptor.Name = name
            oCols = oTableDescriptor.getColumns()

            
            
            for i in range(0, len(cols)):
                if len(dt) > i:
                    oCol = oCols.createDataDescriptor()
                    oCol.Name = cols[i]
                    if i == 0:
                        oCol.Description = "Primary Key"
                    if dt[i] == "BIT":
                        oCol.Type = BIT
                        oCol.Precision = 2
                    elif dt[i] == "TINYINT":
                        oCol.Type = TINYINT
                        oCol.Precision = 4
                    elif dt[i] == "SMALLINT":
                        oCol.Type = SMALLINT
                        oCol.Precision = 8
                    elif dt[i] == "INTEGER":
                        oCol.Type = INTEGER
                        oCol.Precision = 16
                    elif dt[i] == "BIGINT":
                        oCol.Type = BIGINT
                        oCol.Precision = 32
                    elif dt[i] == "FLOAT":
                        oCol.Type = FLOAT
                        oCol.Precision = 64
                    elif dt[i] == "REAL":
                        oCol.Type = REAL
                    elif dt[i] == "DOUBLE":
                        oCol.Type = DOUBLE
                        oCol.Precision = 128
                    elif dt[i] == "NUMERIC":
                        oCol.Type = NUMERIC
                    elif dt[i] == "DECIMAL":
                        oCol.Type = DECIMAL
                    elif dt[i] == "CHAR":
                        oCol.Type = CHAR
                        oCol.Precision = 128
                    elif dt[i] == "VARCHAR":
                        oCol.Type = VARCHAR
                        oCol.Precision = 256
                    elif dt[i] == "LONGVARCHAR":
                        oCol.Type = LONGVARCHAR
                        oCol.Precision = 512
                    elif dt[i] == "DATE":
                        oCol.Type = DATE
                    elif dt[i] == "TIME":
                        oCol.Type = TIME
                    elif dt[i] == "TIMESTAMP":
                        oCol.Type = TIMESTAMP
                    elif dt[i] == "BINARY":
                        oCol.Type = BINARY
                    elif dt[i] == "VARBINARY":
                        oCol.Type = VARBINARY
                    elif dt[i] == "LONGVARBINARY":
                        oCol.Type = LONGVARBINARY
                        oCol.Precision = 2147483647
                    elif dt[i] == "SQLNULL":
                        oCol.Type = SQLNULL
                    elif dt[i] == "OTHER":
                        oCol.Type = OTHER
                    oCols.appendByDescriptor(oCol)
            
            
            
        
            oTables.appendByDescriptor(oTableDescriptor)

            
            
            
            return True
          except:
            return False
          
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean RemoveTable(in string name, in string con)
    def RemoveTable(self, name, con):
        if OOoRTC.base_comp.ConnectionList.has_key(con):
          try:
            oTables = OOoRTC.base_comp.ConnectionList[con]["Connection"].getTables()
            oDBTables = oTables.createEnumeration()

            while oDBTables.hasMoreElements():
                oTable = oDBTables.nextElement()
                if str(oTable.Name) == name:
                    oTables.dropByName(name)
                    return True
            return False
          except:
            return False
          
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean AddDataBase(in string name)
    def AddDataBase(self, name):
        try:
            names = OOoRTC.base_comp.base._context.getElementNames()
            for i in names:
                if name == str(i):
                    return False
            
            dbURL = "C:\\Users\\Nobuhiko\\Desktop\\OOoBaseRTC\\" + name + ".odb"
            ofile= os.path.abspath(dbURL)

            oDB = OOoRTC.base_comp.base._context.createInstance()
            oDB.URL = "sdbc:embedded:hsqldb"

            p = PropertyValue()
            properties = (p,)

            oDB.DatabaseDocument.storeAsURL(ofile, properties)

            oDS = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(unohelper.systemPathToFileUrl(ofile),"_blank", 0, () )
            OOoRTC.base_comp.base._context.registerObject(name,oDS.DataSource)
            oDS.close(True)

            return True
        except:
            return False
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # boolean RemoveDataBase(in string name)
    def RemoveDataBase(self, name):
        
        try:
            OOoRTC.base_comp.base._context.revokeObject(name)
            return True
        except:
            return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    


##
# OpenOffice Baseを操作するためのRTCのクラス
##

class OOoBaseControl(OpenRTM_aist.DataFlowComponentBase):
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    
    
    
    self.ConnectionList = {}
    self.ResultSet = {}

    try:
      self.base = OOoBase()
    except NotOOoBaseException:
      return

    self._DataBasePort = OpenRTM_aist.CorbaPort("DataBase")
    self._database = mDataBase_i()

    

    
    
    return

  ##
  # 実行周期を設定する関数
  ##
  def m_setRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # 活性化するための関数
  ## 
  def m_activate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # 不活性化するための関数
  ##
  def m_deactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  


  ##
  # 初期化処理用コールバック関数
  ##
  def onInitialize(self):
    
    OOoRTC.base_comp = self

    self._DataBasePort.registerProvider("database", "DataBase::mDataBase", self._database)
    self.addPort(self._DataBasePort)

    
    
    
    return RTC.RTC_OK

  ##
  # 非活性化処理用コールバック関数
  ##
  
  def onDeactivated(self, ec_id):
    for i,j in self.ConnectionList.items():
      j["Statement"].close()
      j["Connection"].close()
      j["Connection"].dispose()
    self.ConnectionList = {}
    self.ResultSet = {}
    return RTC.RTC_OK
  
  def onExecute(self, ec_id):
    
    

    return RTC.RTC_OK

  ##
  # 終了処理用コールバック関数
  ##
  
  def on_shutdown(self, ec_id):
      OOoRTC.base_comp = None
      return RTC.RTC_OK



##
# コンポーネントを活性化してBaseの操作を開始する関数
##

def Start():
    
    if OOoRTC.base_comp:
        OOoRTC.base_comp.m_activate()

##
# コンポーネントを不活性化してBaseの操作を終了する関数
##

def Stop():
    
    if OOoRTC.base_comp:
        OOoRTC.base_comp.m_deactivate()


##
# コンポーネントの実行周期を設定する関数
##

def Set_Rate():
    pass
    """if OOoRTC.base_comp:
      try:
        writer = OOoDraw()
      except NotOOoWriterException:
          return

      oWriterPages = draw.drawpages
      for i in range(0, oDrawPages.Count):
        oDrawPage = oDrawPages.getByIndex(i)
        forms = oDrawPage.getForms()
        for j in range(0, forms.Count):
          st_control = oDrawPage.getForms().getByIndex(j).getByName('Rate')
          if st_control:
            try:
              text = float(st_control.Text)
            except:
               return
              
            OOoRTC.draw_comp.m_setRate(text)"""
      
      

      
        
        
      
      

      



      
  



##
#RTCをマネージャに登録する関数
##
def OOoBaseControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooobasecontrol_spec)
  manager.registerFactory(profile,
                          OOoBaseControl,
                          OpenRTM_aist.Delete)


def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoBaseControlInit(manager)

  
  comp = manager.createComponent(imp_id)






          

##
# RTC起動の関数
##

def createOOoBaseComp():
                        
    
    if OOoRTC.mgr == None:
      OOoRTC.mgr = OpenRTM_aist.Manager.init(['OOoBase.py'])
      OOoRTC.mgr.setModuleInitProc(MyModuleInit)
      OOoRTC.mgr.activateManager()
      OOoRTC.mgr.runManager(True)
    else:
      MyModuleInit(OOoRTC.mgr)
      
          

    try:
      base = OOoBase()
    except NotOOoBaseException:
      return

    
    MyMsgBox('',u'RTCを起動しました')


    
    
    return None




##
# メッセージボックス表示の関数
# title：ウインドウのタイトル
# message：表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)



##
# OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
    msgbox.execute()
    msgbox.dispose()




##
# OpenOffice Baseを操作するためのクラス
##

class OOoBase():
  def __init__(self):
    self._ctx = XSCRIPTCONTEXT.getComponentContext()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._context = self._ctx.ServiceManager.createInstanceWithContext('com.sun.star.sdb.DatabaseContext', self._ctx)
    
    self._oRowSet = self._ctx.ServiceManager.createInstanceWithContext('com.sun.star.sdb.RowSet', self._ctx)
  




g_exportedScripts = (createOOoBaseComp, Start, Stop, Set_Rate)
