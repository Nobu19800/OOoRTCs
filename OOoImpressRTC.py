# -*- coding: utf-8 -*-



##
#   @file OOoImpressRTC.py
#   @brief OOoImpressControl Component

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math

from os.path import expanduser
sv = sys.version_info

if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']


import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming





import uno
import unohelper
from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext

from com.sun.star.awt.FontWeight import BOLD 
from com.sun.star.awt import Point, Size





import OOoRTC
from ImpressControl import *




#comp_num = random.randint(1,3000)
imp_id = "OOoImpressControl"# + str(comp_num)







oooimpresscontrol_spec = ["implementation_id", imp_id,
                  "type_name",         imp_id,
                  "description",       "Openoffice Impress Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.SlideNumberInRelative", "1",
                  "conf.default.SlideFileInitialNumber", "0",
                  "conf.__widget__.SlideNumberInRelative", "radio",
                  "conf.__widget__.SlideFileInitialNumber", "spin",
                  "conf.__constraints__.SlideNumberInRelative", "(0,1)",
                  "conf.__constraints__.SlideFileInitialNumber", "0<=x<=1000",
                  ""]





    
  


##
# @class OOoImpressControl
# @brief OpenOffice Impressを操作するためのRTCのクラス
#

class OOoImpressControl(ImpressControl):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    ImpressControl.__init__(self, manager)

    
    try:
      self.impress = OOoImpress()
    except NotOOoWtiterException:
      return


    
    
    return

  

  
  ##
  # @brief 不活性化時のコールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onDeactivated(self, ec_id):
    ImpressControl.onDeactivated(self, ec_id)
    Presentation = self.impress.document.Presentation
    Presentation.end()


    
    return RTC.RTC_OK

  ##
  # @brief 活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onActivated(self, ec_id):
    ImpressControl.onActivated(self, ec_id)
    Presentation = self.impress.document.Presentation
    Presentation.UsePen = True
    Presentation.start()

    

    
    mflag = True
    while mflag:
      try:
        pagenum = self.SlideFileInitialNumber[0]
        if pagenum < 0:
          pagenum = 0
        if Presentation.Controller.getSlideCount() <= pagenum:
          pagenum = Presentation.Controller.getSlideCount() - 1
        Presentation.Controller.gotoSlideIndex(pagenum)
        mflag = False
      except:
        pass
    
    return RTC.RTC_OK
  
  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  def onInitialize(self):
    ImpressControl.onInitialize(self)
    OOoRTC.impress_comp = self

    
    
    return RTC.RTC_OK



  ##
  # @brief スライド番号の進める、戻す
  # @param self 
  # @param num 進めるスライド数
  def changeSlideNum(self, num):
    Presentation = self.impress.document.Presentation
    if self.SlideNumberInRelative[0] == 0:
        pagenum = num
    else:
        pagenum = Presentation.Controller.getCurrentSlideIndex() + num
    if pagenum < 0:
        pagenum = 0
    if Presentation.Controller.getSlideCount() <= pagenum:
        pagenum = Presentation.Controller.getSlideCount() - 1
    if pagenum != Presentation.Controller.getCurrentSlideIndex():
        Presentation.Controller.gotoSlideIndex(pagenum)
    


  ##
  # @brief アニメーションを進める、戻す
  # @param self 
  # @param num 進めるアニメーションの数
  def changeEffeceNum(self, num):
    Presentation = self.impress.document.Presentation
    if num > 0:
        for i in range(0, num):
          Presentation.Controller.getSlideShow().nextEffect()
    elif num < 0:
        for i in range(0, -num):
          Presentation.Controller.getSlideShow().previousEffect()


  ##
  # @brief スライド番号取得
  # @param self 
  # @return スライド番号
  def getSlideNum(self):
    Presentation = self.impress.document.Presentation
          
    return Presentation.Controller.getCurrentSlideIndex()
  

  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    ImpressControl.onExecute(self, ec_id)  
    
        

    return RTC.RTC_OK

  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  
  def onFinalize(self):
      ImpressControl.onFinalize(self)
      OOoRTC.impress_comp = None
      return RTC.RTC_OK



##
# @brief コンポーネントを活性化してImpressの操作を開始する関数
#

def Start():
    
    if OOoRTC.impress_comp:
        OOoRTC.impress_comp.mActivate()

##
# @brief コンポーネントを不活性化してImpressの操作を終了する関数
#

def Stop():
    
    if OOoRTC.impress_comp:
        OOoRTC.impress_comp.mDeactivate()


##
# @brief コンポーネントの実行周期を設定する関数
#

def Set_Rate():
    pass

      
      

      
        
        
      
      

      



      
  



##
# @brief RTCをマネージャに登録する関数
# @param manager マネージャーオブジェクト
def OOoImpressControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=oooimpresscontrol_spec)
  manager.registerFactory(profile,
                          OOoImpressControl,
                          OpenRTM_aist.Delete)


##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoImpressControlInit(manager)

  
  comp = manager.createComponent(imp_id)






          

##
# @brief RTC起動の関数
#

def createOOoImpressComp():
    if OOoRTC.impress_comp:
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
      impress = OOoImpress()
    except NotOOoImpressException:
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
# @brief OpenOffice Impressを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部を改変
##

class OOoImpress(Bridge):
  def __init__(self):
    Bridge.__init__(self)
    if not self._document.supportsService('com.sun.star.presentation.PresentationDocument'):
      self.run_errordialog(title='エラー', message='このマクロはOpenOffice.org Impressの中で実行してください')
      raise NotOOoImpressException()
    self.__current_controller = self._document.CurrentController
    self.__drawpages = self._document.DrawPages

  @property
  def drawpages(self): return self.__drawpages
  @property
  def document(self): return self._document
  



    

    

    
    


g_exportedScripts = ( createOOoImpressComp, Start, Stop, Set_Rate)
