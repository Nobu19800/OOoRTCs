# -*- coding: utf-8 -*-


##
#   @file OOoBaseRTC.py
#   @brief OOoBaseControl Component

import optparse
import sys,os,platform
import codecs
import re
import time
import random
import commands
import math


from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', home+'/OOoRTC/BaseIDL', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\BaseIDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']
    
    


import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming

import threading



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




user_profile = os.environ['USERPROFILE'] + "\\Documents"
if os.name == 'posix':
    user_profile = '~'


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
                  "conf.default.filepath", user_profile,
                  "conf.__widget__.filepath", "text",
                  ""]

##
# @class mDataBase_i
# @brief サービスポートDataBase
#
class mDataBase_i (DataBase__POA.mDataBase):
    
    ##
    # @brief コンストラクタ
    # @param self 
    # @param m_comp OOoBaseRTC
    #
    def __init__(self, m_comp):
        
        self.m_comp = m_comp
        self._mutex = threading.RLock()
        
        
        
        

    ##
    # @brief データベースと接続する関数
    # @param self 
    # @param name データベースの登録名
    # @param usr_name ユーザー名
    # @param passward パスワード
    # @return 成功ならTrue、失敗ならFalse
    #
    def setConnection(self, name, usr_name, passward):
        
        if self.m_comp.ConnectionList.has_key(name):
          return True
        guard = OpenRTM_aist.ScopedLock(self._mutex)
        try:
          tmp = {}
          db = self.m_comp.base._context.getByName(name)
          tmp["Connection"] = db.getConnection(usr_name,passward)
          tmp["Statement"] = tmp["Connection"].createStatement()
          
          self.m_comp.ConnectionList[name] = tmp
          del guard
          return True
        except:
          if guard:
              del guard
          return False
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief データベースに問い合わせて結果セットを取得する関数
    # @param self 
    # @param name 結果セットの名前(設定した名前で辞書オブジェクトに格納されます)
    # @param con データベースの登録名
    # @param oSQL SQL文
    # @return 成功ならTrue、失敗ならFalse
    #
    def executeQuery(self, name, con, oSQL):
        
        if self.m_comp.ConnectionList.has_key(con):
          guard = OpenRTM_aist.ScopedLock(self._mutex)
          try:
            self.m_comp.ResultSet[name] = self.m_comp.ConnectionList[con]["Statement"].executeQuery(oSQL)
            del guard
            return True
          except:
            if guard:
              del guard
            return False
          raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        else:
          return False
        

    ##
    # @brief 次のデータレコードへ移動する関数
    # @param self 
    # @param name 結果セットの名前
    # @return 成功ならTrue、失敗ならFalse
    #
    def ResultSetNext(self, name):
        
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].next()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 前のデータレコードへ移動する
    # @param self 
    # @param name 結果セットの名前
    # @return 成功ならTrue、失敗ならFalse
    #
    def ResultSetPrevious(self, name):
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].previous()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 最初のデータレコードへ移動する
    # @param self 
    # @param name 結果セットの名前
    # @return 成功ならTrue、失敗ならFalse
    #
    def ResultSetFirst(self, name):
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].first()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 最後のデータレコードへ移動する
    # @param self 
    # @param name 結果セットの名前
    # @return 成功ならTrue、失敗ならFalse
    #
    def ResultSetLast(self, name):
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].last()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 最初のデータレコードの前へ移動する
    # @param self 
    # @param name 結果セットの名前
    # @return 成功ならTrue、失敗ならFalse
    # 
    def ResultSetBeforeFirst(self, name):
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].beforeFirst()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 最後のデータレコードの次へ移動する
    # @param self 
    # @param name 結果セットの名前
    # @return 成功ならTrue、失敗ならFalse
    # 
    def ResultSetAfterLast(self, name):
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].afterLast()
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief Byte型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getByte(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return chr(self.m_comp.ResultSet[name].getByte(num))
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief Short型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getShort(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return int(self.m_comp.ResultSet[name].getShort(num))
        return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief Long型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getLong(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return long(self.m_comp.ResultSet[name].getLong(num))
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief Float型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getFloat(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return float(self.m_comp.ResultSet[name].getFloat(num))
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief Double型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getDouble(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return self.m_comp.ResultSet[name].getDouble(num)
        else:
          return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief bool型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getBoolean(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return bool(self.m_comp.ResultSet[name].getBoolean(num))
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief String型のデータを取得する関数
    # @param self 
    # @param name 結果セットの名前
    # @param num 列番号
    # @return 取得したデータ
    #
    def getString(self, name, num):
        if self.m_comp.ResultSet.has_key(name):
          return str(self.m_comp.ResultSet[name].getString(num))
        else:
          return ""
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 登録されたデータベースの名前を取得
    # @param self 
    # @return データベース名のリスト
    #
    def getDataBaseNames(self):
        Ans = []
        guard = OpenRTM_aist.ScopedLock(self._mutex)
        names = self.m_comp.base._context.getElementNames()
        for i in names:
            Ans.append(str(i))
        del guard
        return Ans
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
         

    ##
    # @brief データベースに存在するテーブルの名前を取得
    # @param self 
    # @param con データベースの登録名
    # @return テーブル名のリスト
    #
    def getDataTableNames(self, con):
        Ans = []
        
        
        
        if self.m_comp.ConnectionList.has_key(con):
            guard = OpenRTM_aist.ScopedLock(self._mutex)
            try:
		
                oDBTables = self.m_comp.ConnectionList[con]["Connection"].getTables().createEnumeration()
                while oDBTables.hasMoreElements():
                    oTable = oDBTables.nextElement()
                    Ans.append(str(oTable.Name))
                del guard
            except:
                if guard:
                  del guard
                Ans.append("ERROR")
                

        return Ans
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief データベースの更新
    # @param self 
    # @param con データベースの登録名
    # @param oSQL SQL文
    # @return 成功ならTrue、失敗ならFalse
    #
    def executeUpdate(self, con, oSQL):
        guard = OpenRTM_aist.ScopedLock(self._mutex)
        oRstDataSources = self.m_comp.base._context.getByName(con)
        
        
        if self.m_comp.ConnectionList.has_key(con):
          try:

            self.m_comp.ConnectionList[con]["Statement"].executeUpdate(oSQL)
            del guard
            return True
          except:
            if guard:
              del guard
            return False
          raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        else:
          del guard
          return False
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief カーソルの行番号を取得
    # @param self 
    # @param name 結果セットの名前
    # @return 行番号
    #
    def getRow(self, name):
        if self.m_comp.ResultSet.has_key(name):
          return int(self.m_comp.ResultSet[name].getRow())
        return 0
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief データベースにテーブルを追加
    # @param self 
    # @param name 追加するテーブル名
    # @param con データベースの登録名
    # @param cols フィールド名のリスト
    # @param dt フィールドのタイプのリスト
    # @return 成功ならTrue、失敗ならFalse
    #
    def AddTable(self, name, con, cols, dt):
        if self.m_comp.ConnectionList.has_key(con):
          guard = OpenRTM_aist.ScopedLock(self._mutex)
          try:
            oTables = self.m_comp.ConnectionList[con]["Connection"].getTables()
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
                        oCol.Precision = 1
                    elif dt[i] == "TINYINT":
                        oCol.Type = TINYINT
                        oCol.Precision = 3
                    elif dt[i] == "SMALLINT":
                        oCol.Type = SMALLINT
                        oCol.Precision = 5
                    elif dt[i] == "INTEGER":
                        oCol.Type = INTEGER
                        oCol.Precision = 10
                    elif dt[i] == "BIGINT":
                        oCol.Type = BIGINT
                        oCol.Precision = 19
                    elif dt[i] == "FLOAT":
                        oCol.Type = FLOAT
                        oCol.Precision = 17
                    elif dt[i] == "REAL":
                        oCol.Type = REAL
                        oCol.Precision = 17
                    elif dt[i] == "DOUBLE":
                        oCol.Type = DOUBLE
                        oCol.Precision = 17
                    elif dt[i] == "NUMERIC":
                        oCol.Type = NUMERIC
                        oCol.Precision = 10
                    elif dt[i] == "DECIMAL":
                        oCol.Type = DECIMAL
                        oCol.Precision = 10
                    elif dt[i] == "CHAR":
                        oCol.Type = CHAR
                        oCol.Precision = 256
                    elif dt[i] == "VARCHAR":
                        oCol.Type = VARCHAR
                        oCol.Precision = 256
                    elif dt[i] == "LONGVARCHAR":
                        oCol.Type = LONGVARCHAR
                        oCol.Precision = 2147483647
                    elif dt[i] == "DATE":
                        oCol.Type = DATE
                    elif dt[i] == "TIME":
                        oCol.Type = TIME
                    elif dt[i] == "TIMESTAMP":
                        oCol.Type = TIMESTAMP
                    elif dt[i] == "BINARY":
                        oCol.Type = BINARY
                        oCol.Precision = 2147483647
                    elif dt[i] == "VARBINARY":
                        oCol.Type = VARBINARY
                        oCol.Precision = 2147483647
                    elif dt[i] == "LONGVARBINARY":
                        oCol.Type = LONGVARBINARY
                        oCol.Precision = 2147483647
                    elif dt[i] == "SQLNULL":
                        oCol.Type = SQLNULL
                    elif dt[i] == "OTHER":
                        oCol.Type = OTHER
                        oCol.Precision = 2147483647
                    oCols.appendByDescriptor(oCol)
            
            
            
        
            oTables.appendByDescriptor(oTableDescriptor)

            
            
            
            del guard
            return True
          except:
            if guard:
              del guard
            return False
          
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    
    ##
    # @brief データベースからテーブルを削除
    # @param self 
    # @param name 削除するテーブル名
    # @param con データベースの登録名
    # @return 成功ならTrue、失敗ならFalse
    #
    def RemoveTable(self, name, con):
        if self.m_comp.ConnectionList.has_key(con):
          guard = OpenRTM_aist.ScopedLock(self._mutex)
          try:
            oTables = self.m_comp.ConnectionList[con]["Connection"].getTables()
            oDBTables = oTables.createEnumeration()

            while oDBTables.hasMoreElements():
                oTable = oDBTables.nextElement()
                if str(oTable.Name) == name:
                    oTables.dropByName(name)
                    del guard
                    return True
            del guard
            return False
          except:
            if guard:
              del guard
            return False
          
        else:
          return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
         

    ##
    # @brief データベースの追加
    # @param self 
    # @param name 追加するデータベース名
    # @return 成功ならTrue、失敗ならFalse
    #
    def AddDataBase(self, name):
        guard = OpenRTM_aist.ScopedLock(self._mutex)
        try:
            
            names = self.m_comp.base._context.getElementNames()
            for i in names:
                if name == str(i):
                    return False
            
            dbURL = self.m_comp.filepath[0] + "\\" + name + ".odb"
            if os.name == 'posix':
                dbURL = self.m_comp.filepath[0] + "/" + name + ".odb"
            ofile= os.path.abspath(dbURL)

            oDB = self.m_comp.base._context.createInstance()
            oDB.URL = "sdbc:embedded:hsqldb"

            p = PropertyValue()
            properties = (p,)

            oDB.DatabaseDocument.storeAsURL(ofile, properties)

            oDS = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(unohelper.systemPathToFileUrl(ofile),"_blank", 0, () )
            self.m_comp.base._context.registerObject(name,oDS.DataSource)
            oDS.close(True)
            del guard
            return True
        except:
            if guard:
              del guard
            return False
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief データベースの削除
    # @param self 
    # @param name 削除するデータベース名
    # @return 成功ならTrue、失敗ならFalse
    #
    def RemoveDataBase(self, name):
        guard = OpenRTM_aist.ScopedLock(self._mutex)
        try:
            self.m_comp.base._context.revokeObject(name)
            del guard
            return True
        except:
            if guard:
              del guard
            return False
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)



        

    



##
# @class OOoBaseControl
# @brief OpenOffice Baseを操作するためのRTCのクラス
#

class OOoBaseControl(OpenRTM_aist.DataFlowComponentBase):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    
    
    
    self.ConnectionList = {}
    self.ResultSet = {}

    try:
      self.base = OOoBase()
    except NotOOoBaseException:
      return

    self._DataBasePort = OpenRTM_aist.CorbaPort("DataBase")
    self._database = mDataBase_i(self)

    global user_profile
    self.filepath = [user_profile]

    

    
    
    return

  ##
  # @brief 実行周期を設定する関数
  # @param self 
  # @param rate 実行周期
  #
  def mSetRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # @brief 活性化するための関数
  # @param self 
  #
  def mActivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # @brief 不活性化するための関数
  # @param self
  #
  def mDeactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  


  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  def onInitialize(self):
    
    OOoRTC.base_comp = self

    self._DataBasePort.registerProvider("database", "DataBase::mDataBase", self._database)
    self.addPort(self._DataBasePort)


    global user_profile
    self.bindParameter("filepath", self.filepath, user_profile)
    
    
    return RTC.RTC_OK

  ##
  # @brief 非活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onDeactivated(self, ec_id):
    for i,j in self.ConnectionList.items():
      j["Statement"].close()
      j["Connection"].close()
      j["Connection"].dispose()
    self.ConnectionList = {}
    self.ResultSet = {}
    return RTC.RTC_OK


  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    
    

    return RTC.RTC_OK

  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  
  def onFinalize(self):
      OOoRTC.base_comp = None
      return RTC.RTC_OK



##
# @brief コンポーネントを活性化してBaseの操作を開始する関数
#

def Start():
    
    if OOoRTC.base_comp:
        OOoRTC.base_comp.mActivate()

##
# @brief コンポーネントを不活性化してBaseの操作を終了する関数
#

def Stop():
    
    if OOoRTC.base_comp:
        OOoRTC.base_comp.mDeactivate()


##
# @brief コンポーネントの実行周期を設定する関数
#

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
              
            OOoRTC.draw_comp.mSetRate(text)"""
      
      

      
        
        
      
      

      



      
  



##
# @brief RTCをマネージャに登録する関数
# @param manager マネージャーオブジェクト
def OOoBaseControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooobasecontrol_spec)
  manager.registerFactory(profile,
                          OOoBaseControl,
                          OpenRTM_aist.Delete)

##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoBaseControlInit(manager)

  
  comp = manager.createComponent(imp_id)






          

##
# @brief RTC起動の関数
#

def createOOoBaseComp():
    if OOoRTC.base_comp:
        MyMsgBox('',OOoRTC.SetCoding('RTCは起動済みです','utf-8'))
        return                         
    
    if OOoRTC.mgr == None:
        if os.name == 'posix':
            home = expanduser("~")
            OOoRTC.mgr = OpenRTM_aist.Manager.init([os.path.abspath(__file__), '-f', home+'/OOoRTC/rtc.conf'])
        elif os.name == 'nt':
            OOoRTC.mgr = OpenRTM_aist.Manager.init([os.path.abspath(__file__), '-f', '.\\rtc.conf'])
        else:
            return

      
        OOoRTC.mgr.setModuleInitProc(MyModuleInit)
        OOoRTC.mgr.activateManager()
        OOoRTC.mgr.runManager(True)
    else:
        MyModuleInit(OOoRTC.mgr)
      
          

    try:
      base = OOoBase()
    except NotOOoBaseException:
      return

    
    MyMsgBox('',OOoRTC.SetCoding('RTCを起動しました','utf-8'))


    
    
    return None




##
# @brief メッセージボックス表示の関数
# @param title ウインドウのタイトル
# @param message 表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)



##
# @brief OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    try:
        msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
        msgbox.execute()
        msgbox.dispose()
    except:
        msgbox = self._toolkit.createMessageBox(self._window,'infobox',1,title,message)
        msgbox.execute()
        msgbox.dispose()




##
# @brief OpenOffice Baseを操作するためのクラス
#

class OOoBase():
  def __init__(self):
    self._ctx = XSCRIPTCONTEXT.getComponentContext()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._context = self._ctx.ServiceManager.createInstanceWithContext('com.sun.star.sdb.DatabaseContext', self._ctx)
    
    self._oRowSet = self._ctx.ServiceManager.createInstanceWithContext('com.sun.star.sdb.RowSet', self._ctx)
  




g_exportedScripts = (createOOoBaseComp, Start, Stop, Set_Rate)
