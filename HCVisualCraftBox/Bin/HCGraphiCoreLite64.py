# Windows - python3
# -*- coding: utf-8 -*-

# 系统模块
import os
import sys
import time
import ctypes

# 第三方模块
import gdsfactory # type: ignore

# Powered By "阳菜"
__version__ = "1.1.0.8"

# 调试开关
DbgPrint = True

# 定义常量
EVENT_ALL_ACCESS = 0x1F0003
INFINITE = 0xFFFFFFFF

# 将得到的字符串参数传入并解析
def main(m_CurrentDirectory, m_CStr0, m_CStr1, m_CStr2, m_CStr3, m_CStr4, m_CStr5, m_CStr6, m_CStr7, m_CStr8):
    # 返回的值的第一个位指明是哪个函数
    match m_CStr0:
        case "0":
            # RGC 矩形耦合光栅
            if (DbgPrint): OutputDebugString(b"Func0---------------------------Run------------------------------")
            # 生成GDS
            c = gdsfactory.Component('straight')
            c << gdsfactory.components.straight(length = float(m_CStr1), width = float(m_CStr2))
            c.write_gds(os.path.dirname(m_CurrentDirectory) + "\\" + m_CStr3 + ".gds")
        case "1":
            # MRR 微环
            if (DbgPrint): OutputDebugString(b"Func1---------------------------Run------------------------------")
            # 生成GDS
            c = gdsfactory.Component('MRR')
            c << gdsfactory.components.ring_single(gap = float(m_CStr1), radius = float(m_CStr2), length_x = float(m_CStr3), 
                                          length_y = float(m_CStr4), cross_section = gdsfactory.cross_section.strip(width = float(m_CStr5)))
            c.write_gds(os.path.dirname(m_CurrentDirectory) + "\\" + m_CStr6 + ".gds")

        case "2":
            # MZI 马赫曾德干涉仪
            if (DbgPrint): OutputDebugString(b"Func2---------------------------Run------------------------------")
            # 生成GDS
            c = gdsfactory.Component("mzi")
            x1 = gdsfactory.cross_section.strip(width = float(m_CStr1))
            mmi = gdsfactory.components.mmi1x2(width = float(m_CStr2), width_taper = float(m_CStr3), width_mmi = float(m_CStr4))
            c << gdsfactory.components.mzi(delta_length = float(m_CStr5), length_y = float(m_CStr6), 
                                                 length_x = float(m_CStr7) , splitter = mmi, cross_section = x1)
            c.write_gds(os.path.dirname(m_CurrentDirectory) + "\\" + m_CStr8 + ".gds")

        case "3":
            # WG  矩形光栅
            if (DbgPrint): OutputDebugString(b"Func3---------------------------Run------------------------------")
            # 生成GDS
            c = gdsfactory.Component('WG')
            c << gdsfactory.components.grating_coupler_rectangular(n_periods = int(m_CStr1), period = float(m_CStr2), 
                                                                         fill_factor = float(m_CStr3), width_grating = float(m_CStr4), length_taper = float(m_CStr5))
            c.write_gds(os.path.dirname(m_CurrentDirectory) + "\\" + m_CStr6 + ".gds")
            
        case _:
            # 都不是的情况
            if (DbgPrint): OutputDebugString(b"FuncErr---------------------------Run------------------------------")
            pass

# 检查凭证 
def CredentialCheck():
    # 如果非HCVisualCraftBox.exe程序创建则退出
    if len(sys.argv) < 2:
        sys.exit(-1)
    if sys.argv[1] != "yangcai666":
        sys.exit(-2)
    return 0

# 获取当前目录路径
def GetCurrentPath():
    if getattr(sys, "frozen", False):
        # 如果当前是exe
        m_CurrentDirectoryPath = os.path.dirname(sys.executable)
    else:
        # 如果当前是Python环境
        m_CurrentDirectoryPath = sys.path[0]
    return m_CurrentDirectoryPath

# 加载需要使用的模块
def LoadModule():
    AddressList = []
    # 加载 kernel32.dll | Connector64.dll
    kernel32 = ctypes.WinDLL("kernel32")
    hConnector = ctypes.WinDLL(os.path.join(CurrentDirectory, "Connector64.dll"))
    AddressList.append(kernel32)        # 0
    AddressList.append(hConnector)      # 1
    return AddressList

# 打开HCVisualCraftBox.exe事件
def WaitSingle():
    while True:
        time.sleep(1)
        if (DbgPrint): OutputDebugString(b"Wait..................")
        # 打开ClickCommandButton事件对象
        event_handle = OpenEvent(EVENT_ALL_ACCESS, False, b"ClickCommandButton")
        if not event_handle:
            continue
        return event_handle
    
if __name__ == "__main__":
    # 检查凭证
    CredentialCheck()
    # 获取当前目录路径
    CurrentDirectory = GetCurrentPath()
    # 加载需要使用的模块
    AddressOfFunction = LoadModule()
    # 声明需要用到的函数原型
    if True:
        # 声明函数原型OutputDebugString
        OutputDebugString = AddressOfFunction[0].OutputDebugStringA
        OutputDebugString.argtypes = [ctypes.c_char_p]
        OutputDebugString.restype = None
        # 声明函数原型OpenEventA
        OpenEvent = AddressOfFunction[0].OpenEventA
        OpenEvent.argtypes = [ctypes.wintypes.DWORD, ctypes.wintypes.BOOL, ctypes.wintypes.LPCSTR]
        OpenEvent.restype = ctypes.wintypes.HANDLE
        # 声明函数原型WaitForSingleObject
        WaitForSingleObject = AddressOfFunction[0].WaitForSingleObject
        WaitForSingleObject.argtypes = [ctypes.wintypes.HANDLE, ctypes.wintypes.DWORD]
        WaitForSingleObject.restype = ctypes.wintypes.DWORD
        # 声明函数原型GetDataByFileMapping
        GetDataByFileMapping = AddressOfFunction[1].GetDataByFileMapping
        GetDataByFileMapping.argtypes = [ctypes.c_char_p, ctypes.c_char_p]
        GetDataByFileMapping.restype = ctypes.c_int
    if (DbgPrint): OutputDebugString(b"111111111111111111111111111111111111111111111111111111111111111")
    # 打开HCVisualCraftBox.exe事件 函数返回则证明打开成功
    hEventRet = WaitSingle()
    if (DbgPrint): OutputDebugString(b"222222222222222222222222222222222222222222222222222222222222222222222")
    pszSymbol = ctypes.c_char_p(b"MemShared")  # 共享内存符号"MemShared"
    szRet = ctypes.create_string_buffer(0x100)  # 为 szRet 分配缓冲区
    while True:
        if (DbgPrint): OutputDebugString(b"33333333333333333333333333333333333333333333333333333333333333")
        WaitForSingleObject(hEventRet, INFINITE)
        if (DbgPrint): OutputDebugString(b"444444444444444444444444444444444444444444444444444444444444444")
        # 得到从界面获取的值并拆分数据
        GetDataByFileMapping(pszSymbol, szRet)
        python_string = szRet.value.decode("ascii")
        Flag = 0
        CStr0 = python_string[0]
        CStr1 = ""
        CStr2 = ""
        CStr3 = ""
        CStr4 = ""
        CStr5 = ""
        CStr6 = ""
        CStr7 = ""
        CStr8 = ""
        for i in python_string[2:]:
            if i == "|": 
                Flag = Flag + 1
                continue
            match Flag:
                case 0: CStr1 =  CStr1 + i
                case 1: CStr2 =  CStr2 + i
                case 2: CStr3 =  CStr3 + i
                case 3: CStr4 =  CStr4 + i
                case 4: CStr5 =  CStr5 + i
                case 5: CStr6 =  CStr6 + i
                case 6: CStr7 =  CStr7 + i
                case 7: CStr8 =  CStr8 + i
        # 执行主函数分析数据并生成文件
        main(CurrentDirectory, CStr0, CStr1, CStr2, CStr3, CStr4, CStr5, CStr6, CStr7, CStr8)