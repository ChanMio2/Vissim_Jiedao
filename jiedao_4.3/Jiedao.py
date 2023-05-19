
# COM-Server
import win32com.client as com


# ==========导入文件==========
Vissim = com.gencache.EnsureDispatch("Vissim.Vissim.430")
Vissim.LoadNet("E:\Desktop\github\jiedao_4.3\jiedao.inp")


# ==========仿真参数==========
Vissim.Simulation.Period = 3600
Vissim.Simulation.SetAttValue('RandomSeed', 42)


# ==========函数==========
# 获取借道左转信号灯组状态
# 1 = 红灯, 2 = 红灯/黄灯, 3 = 绿灯, 4 = 黄灯
def func_sigstate(SC_number, SG_number):
    State_of_SG = Vissim.Net.SignalControllers(SC_number).SignalGroups.GetSignalGroupByNumber(SG_number).AttValue('STATE')
    return State_of_SG

# 设置车道禁行车辆类型
def func_blockedveh(link, lane, vehclass, type = True):
    Vissim.Net.Links.GetLinkByNumber(link).SetAttValue2('LANECLOSED', lane, vehclass, type)


# ==========仿真==========
try:
    for sim_step in range(1, 36000):
        Vissim.Simulation.RunSingleStep()
        print(f'借道左转信号灯状态:{func_sigstate(1, 1)}')
        if func_sigstate(1, 1) != 3:
            func_blockedveh(1, 2, 10)
            func_blockedveh(10002, 2, 10)
        else:
            func_blockedveh(1, 2, 10)
            func_blockedveh(10002, 2, 10, False)
        sim_step += 1
except:
    Vissim = None
    print('结束')