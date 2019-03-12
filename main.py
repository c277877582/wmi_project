# coding: utf-8

import wmi


from win32com.client import GetObject
import win32api

#################
#wmi = GetObject('winmgmts:/<命名空间>')
#<instance> = wmi.ExecQuery('select * from <顶级类> ')    #此处用的WQL语句，类似SQL语句

#for item in <instance>:
#    <custom_variable> = item.<类属性>
################



# def vmi_main():
#     c = wmi.WMI()
#     for printer in c.CIM_Printer():
#         print("{}                    {}".format(printer.Attributes, printer.Name))

def service_main():
    wmi = GetObject("winmgmts:/root/cimv2")
    products = wmi.ExecQuery("select * from CIM_Product")

    for product in products:
        print("{}====={}====={}".format(product.Caption,product.Description,product.Name))


# def getobj_main():
#     wmi = GetObject("winmgmts:/root/cimv2")
#     printerset = wmi.ExecQuery("select * from CIM_Printer")
#     if printerset.Count == 0:
#         print("No Printers Installed!")
#
#     for printer in printerset:
#         # print("{}                    {}".format(printer.Attributes,printer.Name))
#         try:
#             if printer.Attributes == 2588:
#                 print("开始打印默认页....")
#                 printer.PrintTestPage()
#         except:
#             pass

# def product_main():
#     c = wmi.WMI("CIM_Scanner")
#     for item in c:
#         # print("="*150)
#         # print("标题：{}".format(item.Caption))
#         # print("描述：{}".format(item.Description))
#         # print("名字：{}".format(item.Name))
#         # print("版本：{}".format(item.Version))
#         # print("Vendor：{}".format(item.Vendor))
#         # print("SKUNumber：{}".format(item.SKUNumber))
#         # print("IdentifyingNumber：{}".format(item.IdentifyingNumber))
#         # print("="*150)
#         print(item.Caption)

if __name__ == '__main__':
    # getobj_main()
    # vmi_main()
    # service_main()
    # product_main()


    c = wmi.WMI()
    for os in c.Win32_Process():
        # print(os.ExecutablePath)
        if os.Name == "TIM.exe":
            print("找到Tim.exe文件....")
            tim_dir = os.ExecutablePath
            print("获取路径中....")
            print(tim_dir)
            print("分割路径....")
            import os
            print("返回上一级目录....")
            turnup_dir = os.path.dirname(os.path.dirname(tim_dir))
            print(turnup_dir)
            uninstall = "TIMUninst.exe"
            print("卸载程序路径....")
            new_path = os.path.join(turnup_dir + "\\" + uninstall)
            print(new_path)
            print("运行卸载程序....")
            win32api.ShellExecute(0, 'open', new_path, '', '', 0)
            print("请点击卸载程序....")