#!/use/bin/env python
# _*_ coding:utf-8 _*_
# Author: senliu <senliu@digitgate.cn>

from classexcel import ExcelOpenPyXl
from logging import handlers
import paramiko
import logging
import psutil
import time
import os
from test_factory import test_lftp_5g_100
import tkinter
from tkinter.messagebox import showinfo
from tkinter import *

usb20_write_command = 'sync;time dd if=/dev/zero of=/mnt/usb2.0/tempfile bs=1k count=102400;time sync'
usb20_read_command = 'time dd if=/mnt/usb2.0/tempfile of=/dev/null bs=10k'
usb30_write_command = 'sync;time dd if=/dev/zero of=/mnt/usb3.0/tempfile bs=1k count=102400;time sync'
usb30_read_command = 'time dd if=/mnt/usb3.0/tempfile of=/dev/null bs=10k'

def logger_use(log_path):  # 定义日志
    log_handle_ch = logging.StreamHandler()
    log_file_name = time.strftime('huitie_SN' + sn_num_input + '_%Y%m%d%H%M%S.log', time.localtime(time.time()))
    log_handle_fh = handlers.RotatingFileHandler(log_path + log_file_name, encoding='utf8',
                                                 maxBytes=10**7, backupCount=100)
    formatter = logging.Formatter('[%(asctime)s] [%(message)s]')
    log_handle_ch.setFormatter(formatter)
    log_handle_fh.setFormatter(formatter)
    logger.addHandler(log_handle_ch)
    logger.addHandler(log_handle_fh)


def ssh_connect(host, user, passwd):  # ssh登录设备
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname=host, username=user, password=passwd, timeout=60)
    return ssh


def test_ping():
    row_num = 2  # excel行
    col_num = 3  # excel列
    # ssh登录慧铁设备，ping远端各个server ip
    try:
        ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    except:
        showinfo('警告', '请确认SSH连接是否正常！！！')
        ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)

    for server_ip in server_ip_list:
        ping_cmd = 'ping -c 1 {}'.format(server_ip)  # ping指令
        sin, sout, serr = ssh_huitie.exec_command(ping_cmd, timeout=30, get_pty=True)
        res = sout.read().strip().decode()
        if 'ttl' in res:
            print('>> 网口通断测试：ping ' + server_ip + '  Pass')
            excel_report.write_cell(row_num, col_num, 'Pass')
            time.sleep(1)
        else:
            print('>> 网口通断测试：ping ' + server_ip + '  Fail')
            excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
            time.sleep(1)
        row_num += 1
        excel_report.save_excel()
    ssh_huitie.close()


def test_iperf3():
    row_num = 8
    col_num = 3
    # 登录远端server，打开iperf3
    ssh_server = ssh_connect(server_ip_list[0], server_user, server_password)
    ssh_server.exec_command('iperf3 -s', timeout=30, get_pty=True)

    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    # client端通过iperf,测试10G网口速率
    for server_ip in server_ip_list[0:2]:
        print('>> 网口速率测试：' + server_ip + '  等待30秒')
        iperf_cmd = f'iperf3 -c {server_ip} -t 30'
        sin, sout, serr = ssh_huitie.exec_command(iperf_cmd, timeout=30, get_pty=True)
        bandwidth_list = sout.read().decode().split('\r\n')
        sender_speed_10g = bandwidth_list[-5].split(' ')[12]
        receiver_speed_10g = bandwidth_list[-4].split(' ')[12]
        # print(bandwidth_list[-5].split(' '))
        if float(sender_speed_10g) >= 3 and float(receiver_speed_10g) >= 3:
            print('>> 网口速率测试：' + server_ip + '  发送速率：' + sender_speed_10g +
                        'Gbits/sec，接收速率：' + receiver_speed_10g + ' Gbits/sec  Pass')
            excel_report.write_cell(row_num, col_num, 'Pass')
        else:
            print('>> 网口速率测试：' + server_ip + '  发送速率：' + sender_speed_10g +
                        'Gbits/sec，接收速率：' + receiver_speed_10g + ' Gbits/sec  Fail')
            excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
        row_num += 1
        excel_report.save_excel()

    # client端通过iperf测试1G网口速率
    for server_ip in server_ip_list[2:6]:
        print('>> 网口速率测试：' + server_ip + '  等待30秒')
        iperf_cmd = f'iperf3 -c {server_ip} -t 30'
        sin, sout, serr = ssh_huitie.exec_command(iperf_cmd, timeout=30, get_pty=True)
        bandwidth_list = sout.read().decode().split('\r\n')
        sender_speed_1g = bandwidth_list[-5].split(' ')[13]
        receiver_speed_1g = bandwidth_list[-4].split(' ')[13]
        # print(bandwidth_list[-5].split(' '))
        if float(sender_speed_1g) >= 900 and float(receiver_speed_1g) >= 900:
            print('>> 网口速率测试：' + server_ip + '  发送速率：' + sender_speed_1g +
                        'Mbits/sec，接收速率：' + receiver_speed_1g + ' Mbits/sec  Pass')
            excel_report.write_cell(row_num, col_num, 'Pass')
        else:
            print('>> 网口速率测试：' + server_ip + '  发送速率：' + sender_speed_1g +
                        'Mbits/sec，接收速率：' + receiver_speed_1g + ' Mbits/sec  Fail')
            excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
        row_num += 1
        excel_report.save_excel()
    ssh_server.close()
    ssh_huitie.close()


def test_rtc():
    row_num = 14
    col_num = 3

    # ssh登录慧铁设备，rtc测试
    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    # 查询时间
    sin, sout, serr = ssh_huitie.exec_command('date', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 'UTC' in res:
        print('>> RTC时间测试：查询时间  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
    else:
        print('>> RTC时间测试：查询时间  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
    time.sleep(1)

    # 设置时间
    ssh_huitie.exec_command('date -s "2021-02-08 12:00:00"', timeout=30, get_pty=True)
    sin, sout, serr = ssh_huitie.exec_command('date', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if '12:00:00' in res and '2021' in res and 'Feb' in res:
        print('>> RTC时间测试：设置时间  Pass')
        excel_report.write_cell(row_num+1, col_num, 'Pass')
    else:
        print('>> RTC时间测试：设置时间  Fail')
        excel_report.write_cell(row_num+1, col_num, 'Fail', font_color_str="FF0000")
    time.sleep(1)

    # 写入RTC芯片并读取
    ssh_huitie.exec_command('hwclock -w', timeout=30, get_pty=True)
    time.sleep(2)
    sin, sout, serr = ssh_huitie.exec_command('hwclock', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if '12:00' in res and '2021' in res:
        print('>> RTC时间写入芯片测试：设置时间  Pass')
        excel_report.write_cell(row_num+2, col_num, 'Pass')
    else:
        print('>> RTC时间写入芯片测试：设置时间  Fail')
        print('>> 错误信息：' + res)
        excel_report.write_cell(row_num+2, col_num, 'Fail', font_color_str="FF0000")
    row_num += 1
    excel_report.save_excel()
    ssh_huitie.close()


def test_temp():
    row_num = 17
    col_num = 3
    # ssh登录慧铁设备，温度传感器测试
    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    sin, sout, serr = ssh_huitie.exec_command('cat /sys/class/hwmon/hwmon0/temp1_input', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 1000 < int(res) < 100000:
        print('>> 温度传感器测试：读取温度 ' + res + '  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
    else:
        print('>> 温度传感器测试：当前温度 ' + res + '  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
    excel_report.save_excel()
    ssh_huitie.close()


def test_eeprom():
    row_num = 18
    col_num = 3
    # ssh登录慧铁设备，eeprom测试
    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    sin, sout, serr = ssh_huitie.exec_command('htgw id HTKJ000100', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 'Success' in res:
        print('>> eeprom测试：写入id，' + res + '  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
    else:
        print('>> eeprom测试：写入id，' + res + '  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")

    ssh_huitie.exec_command('reboot', timeout=30, get_pty=True)
    time.sleep(1)
    ssh_huitie.close()
    print('>> eeprom测试：重启设备，等待90秒')
    # for m in range(90, 0, -1):  # 等待启动时间
    print("\r>> 倒计时90秒,请等待......")
    time.sleep(90)
    print("\r>> 倒计时结束......")

    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    for server_ip in server_ip_list:
        ping_cmd = 'ping -c 1 {}'.format(server_ip)  # ping指令
        sin, sout, serr = ssh_huitie.exec_command(ping_cmd, timeout=30, get_pty=True)
        res = sout.read().strip().decode()
        time.sleep(0.1)
        if 'ttl' in res:
            continue
        else:
            print('>> eeprom测试：设备重启后通信异常，' + server_ip + ' ping Fail')
    sin, sout, serr = ssh_huitie.exec_command('htgw id', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 'HTKJ000100' in res:
        print('>> eeprom测试：读取id，' + res + '  Pass')
        excel_report.write_cell(row_num+1, col_num, 'Pass')
    else:
        print('>> eeprom测试：读取id' + res + '  Fail')
        excel_report.write_cell(row_num+1, col_num, 'Fail', font_color_str="FF0000")
    excel_report.save_excel()
    ssh_huitie.close()


def test_usb_m_and_check_ssd():
    row_num = 20
    col_num = 3

    # ssh登录慧铁设备，检测固态硬盘
    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    sin, sout, serr = ssh_huitie.exec_command('fdisk -l', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 'Disk /dev/sda' in res:
        print('>> USB_M口测试和固态硬盘检测：检测固态硬盘  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
    else:
        print('>> USB_M口测试和固态硬盘检测：检测固态硬盘  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
    excel_report.save_excel()

    # USB_M口检测USB 2.0
    while True:
        # input_info = input('请在USB_M接口插入USB 2.0，输入y或者yes继续测试：')
        askback = tkinter.messagebox.askyesno(title='请确认', message='请在USB_M接口插入USB 2.0')
        if askback == FALSE:
            sys.exit()
        print('>> USB_M口测试和固态硬盘检测：开始检测UBS 2.0')
        time.sleep(1)
        break
    sin, sout, serr = ssh_huitie.exec_command('dmesg -c', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 'new high-speed USB device' in res:
        print('>> USB_M口测试和固态硬盘检测：USB_M口检测UBS 2.0  Pass')
        excel_report.write_cell(row_num+1, col_num, 'Pass')
    else:
        print('>> USB_M口测试和固态硬盘检测：USB_M口检测UBS 2.0  Fail')
        excel_report.write_cell(row_num+1, col_num, 'Fail', font_color_str="FF0000")

    # USB 2.0写入速率测试
    print('>> 开始测试USB2.0写入速率......')
    ssh_huitie.exec_command('mkdir /mnt/usb2.0', timeout=30, get_pty=True)
    ssh_huitie.exec_command('mount /dev/sdb /mnt/usb2.0', timeout=30, get_pty=True)
    time.sleep(2)
    sin, sout, serr = ssh_huitie.exec_command(usb20_write_command, timeout=30, get_pty=True)
    res_list = sout.read().decode().split('\r\n')
    print(res_list[2])
    write_speed = res_list[2].split(' ')[9]
    print('>> USB2.0写入速率:', write_speed)
    excel_report.write_cell(row_num+2, col_num, write_speed)
    excel_report.save_excel()

    # USB 2.0读取速率测试
    print('>> 开始测试USB2.0读取速率......')
    ssh_huitie.exec_command('sysctl -w vm.drop_caches=3', timeout=30, get_pty=True)
    time.sleep(2)
    sin, sout, serr = ssh_huitie.exec_command(usb20_read_command, timeout=30, get_pty=True)
    res_list = sout.read().decode()
    read_speed = res_list.split('\r\n')[2].split(' ')[9]
    print('>> USB2.0读取速率:', read_speed)
    ssh_huitie.exec_command('umount /dev/sdb /mnt/usb2.0', timeout=30, get_pty=True)
    ssh_huitie.exec_command('rm /mnt/usb2.0/tempfile', timeout=30, get_pty=True)

    excel_report.write_cell(row_num + 3, col_num, read_speed)
    excel_report.save_excel()

    # USB_M口检测USB 3.0
    while True:
        # input_info = input('请在USB_M接口插入USB 3.0，输入y或者yes继续测试：')
        askback = tkinter.messagebox.askyesno(title='请确认', message='请在USB_M接口插入USB 3.0')
        # if 'y' in input_info or 'yes' in input_info:
        #     print('>> USB_M口测试和固态硬盘检测：开始检测UBS 3.0')
        #     time.sleep(1)
        #     break
        if askback == FALSE:
            sys.exit()
        print('>> USB_M口测试和固态硬盘检测：开始检测UBS 3.0')
        time.sleep(1)
        break
    sin, sout, serr = ssh_huitie.exec_command('dmesg -c', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if 'new SuperSpeed Gen 1 USB device' in res:
        print('>> USB_M口测试和固态硬盘检测：USB_M口检测UBS 3.0  Pass')
        excel_report.write_cell(row_num+4, col_num, 'Pass')
    else:
        print('>> USB_M口测试和固态硬盘检测：USB_M口检测UBS 3.0  Fail')
        excel_report.write_cell(row_num+4, col_num, 'Fail', font_color_str="FF0000")



    # USB 3.0写入速率测试
    print('>> 开始测试USB3.0写入速率......')
    ssh_huitie.exec_command('mkdir /mnt/usb3.0', timeout=30, get_pty=True)
    ssh_huitie.exec_command('mount /dev/sdb /mnt/usb3.0', timeout=30, get_pty=True)
    time.sleep(2)
    sin, sout, serr = ssh_huitie.exec_command(usb30_write_command, timeout=30, get_pty=True)
    res_list = sout.read().decode()
    print(res_list.split('\r\n')[2].split(' '))
    write_speed = res_list.split('\r\n')[2].split(' ')[9]
    print('>> USB3.0写入速率:', write_speed)
    excel_report.write_cell(row_num + 5, col_num, write_speed)
    excel_report.save_excel()

    # USB 3.0读取速率测试
    print('>> 开始测试USB3.0读取速率......')
    ssh_huitie.exec_command('sysctl -w vm.drop_caches=3', timeout=30, get_pty=True)
    time.sleep(2)
    sin, sout, serr = ssh_huitie.exec_command(usb30_read_command, timeout=30, get_pty=True)
    res_list = sout.read().decode()
    read_speed = res_list.split('\r\n')[2].split(' ')[9]
    print('>> USB3.0读取速率:', read_speed)
    ssh_huitie.exec_command('umount /dev/sdb /mnt/usb3.0', timeout=30, get_pty=True)
    ssh_huitie.exec_command('rm /mnt/usb3.0/tempfile', timeout=30, get_pty=True)
    excel_report.write_cell(row_num + 6, col_num, read_speed)

    excel_report.save_excel()
    ssh_huitie.close()



def test_usb_s():
    row_num = 27
    col_num = 3
    disk_nums_before = len(psutil.disk_partitions())  # 检测当前盘符数量
    # ssh登录慧铁设备，虚拟U盘测试
    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
    cmds_config = ['depmod', 'modprobe libcomposite', 'modprobe usb_f_mass_storage',
                   'modprobe g_mass_storage file=/dev/sda removable=1']
    for cmd in cmds_config:
        ssh_huitie.exec_command(cmd, timeout=30, get_pty=True)
    while True:
        # res = input('请将USB线一端接入USB_S口，一端接入测试电脑，并等待电脑检测到U盘后，输入y或者yes继续测试: ')
        askback = tkinter.messagebox.askyesno(title='请确认', message='请将USB线一端接入USB_S口，一端接入测试电脑，并等待电脑检测到U盘')
        # if 'y' in res or 'yes' in res:
        #     print('>> USB_S口虚拟U盘测试：开始检测')
        #     time.sleep(1)
        #     break
        if askback == FALSE:
            sys.exit()
        print('>> USB_S口虚拟U盘测试：开始检测')
        time.sleep(1)
        break
    disk_nums_after = len(psutil.disk_partitions())
    if disk_nums_after == disk_nums_before + 1:
        print('>> USB_S口虚拟U盘测试：检测虚拟U盘  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
    else:
        print('>> USB_S口虚拟U盘测试：检测虚拟U盘  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
    excel_report.save_excel()
    ssh_huitie.close()


def test_mac():
    row_num = 28
    col_num = 3
    # ssh登录慧铁设备，MAC地址测试
    ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)

    """获取所有eth的MAC地址，并返回列表保存"""
    sin, sout, serr = ssh_huitie.exec_command('ip link', get_pty=True)
    res = sout.read().strip().decode()
    eth_ports = ['eth1', 'eth2', 'eth3', 'eth4', 'eth5', 'eth6']
    eth_macs = []
    for port in eth_ports:
        eth_mac = re.search(f'.*{port}:.*\n.*link/ether (.*) brd.*', res)
        if eth_mac:
            eth_macs.append(eth_mac.group(1))
        else:
            print('>> MAC地址测试：获取MAC失败')
    id_hex = hex(100).replace('0x', '')  # ID号转换为十六进制MAC

    for eth_num in range(0, 6):
        if '48:54:00:00:' + id_hex + ':' in eth_macs[eth_num]:
            print('>> MAC地址测试：eth' + str(eth_num+1) + '端口 MAC地址 ' + eth_macs[eth_num] + '  Pass')
            excel_report.write_cell(row_num, col_num, 'Pass')
            row_num += 1
            time.sleep(0.5)
        else:
            print('>> MAC地址测试：eth' + str(eth_num+1) + '端口 MAC地址 ' + eth_macs[eth_num] + '  Fail')
            excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
            row_num += 1
            time.sleep(0.5)
    # 重启后再次检查date时间
    sin, sout, serr = ssh_huitie.exec_command('date', timeout=30, get_pty=True)
    res = sout.read().strip().decode()
    if '2021' in res and 'Feb' in res and 'UTC' in res:
        print('>> RTC时间测试：重启后查询时间 ' + res + '  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
    else:
        print('>> RTC时间测试：重启后查询时间 ' + res + '  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
    time.sleep(1)
    excel_report.save_excel()
    ssh_huitie.close()


def test_5g():
    row_num = 35
    col_num = 3
    print('>> 5G网口业务测试')
    time.sleep(1)
    # print('>> cmd中执行指令：'
    #             'pytest -s -v -k "100" --html=D:/ht_test/test_lftp_5g_10.html D:/ht_test/huit/test_factory.py')
    print('>> 正在执行test_lftp_5g_100')
    test_lftp_5g_100()
    time.sleep(1)
    # for m in range(120, 0, -1):  # 等待启动时间
    #     print("\r正在倒计时120s, 请等待......")
    #     time.sleep(1)
    # print("\r倒计时结束......")
    # lftp_5g_res = input('请输入测试结果(通过：y ，失败：n)：')
    # if lftp_5g_res == 'y' or lftp_5g_res == 'yes':
    #     print('>> 5G网口业务测试：  Pass')
    #     excel_report.write_cell(row_num, col_num, 'Pass')
    #     excel_report.write_cell(row_num+1, col_num, 'Pass')
    # else:
    #     print('>> 5G网口业务测试：  Fail')
    #     excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
    #     excel_report.write_cell(row_num+1, col_num, 'Fail', font_color_str="FF0000")

    askback = tkinter.messagebox.askyesno(title='请确认', message='请输入测试结果是否PASS')
    if askback == FALSE:
        print('>> 5G网口业务测试：  Fail')
        excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
        excel_report.write_cell(row_num + 1, col_num, 'Fail', font_color_str="FF0000")
    else:
        print('>> 5G网口业务测试：  Pass')
        excel_report.write_cell(row_num, col_num, 'Pass')
        excel_report.write_cell(row_num + 1, col_num, 'Pass')

    excel_report.save_excel()


def test_poe():
    row_num = 37
    col_num = 3
    # ssh登录慧铁设备，MAC地址测试
    ssh_huitie = ssh_connect('192.168.3.101', huitie_user, huitie_password)
    for eth_num in ('1', '2'):
        showinfo(title='请确认', message='请将%s口网线连接负载poe！！！'% eth_num)
        print('>> POE带载测试：开启 eth' + eth_num + ' poe 供电')
        ssh_huitie.exec_command('poectl on eth' + eth_num, timeout=30, get_pty=True)
        showinfo(title='请确认', message='请确认负载状态为ON！！！')
        print('>> POE带载测试：将 eth' + eth_num + ' 网线连接负载，保持30秒')
        # for m in range(30, 0, -1):  # 等待启动时间
        #     print("\r倒计时30s, 请等待......")
        #     time.sleep(1)
        print('>> 倒计时30s, 请等待......')
        time.sleep(30)

        print("\r>> 倒计时结束......")
        print('>> POE带载测试：关闭 eth' + eth_num + ' poe 供电')
        ssh_huitie.exec_command('poectl off eth' + eth_num, timeout=30, get_pty=True)
        time.sleep(1)

        # poe_res = input('请输入 eth' + eth_num+' 测试结果(通过：y ，失败：n)：')
        # if poe_res == 'y' or poe_res == 'yes':
        #     print('>> POE带载测试：  Pass')
        #     excel_report.write_cell(row_num, col_num, 'Pass')
        # else:
        #     print('>> POE带载测试：  Fail')
        #     excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
        showinfo(title='请确认', message='请确认负载状态为OFF！！！')
        askback = tkinter.messagebox.askyesno(title='请确认', message='请输入 eth' + eth_num +' 测试结果是否PASS')
        if askback == FALSE:
            print('>> POE带载测试：  Fail')
            excel_report.write_cell(row_num, col_num, 'Fail', font_color_str="FF0000")
        else:
            print('>> POE带载测试：  Pass')
            excel_report.write_cell(row_num, col_num, 'Pass')

        row_num += 1
        excel_report.save_excel()
    ssh_huitie.close()


def check_test_report():
    # res = []
    for i in range(2, 35):
        if excel_report.read_cell_value(i, 3) == 'Fail' or excel_report.read_cell_value(i, 3) is None:
            print('>> 慧铁1U网关，软件测试  Fail。设备存在异常，请检修！')
            showinfo(title='请确认', message='测试结束， Fail，设备存在异常，请检修！')
            break


# def test_5g_lftp():
#     row_num = 30
#     col_num = 3
#
#     # 登录远端server
#     ssh_server = ssh_connect(server_ip_list[0], server_user, server_password)
#
#     # ssh登录慧铁设备，MAC地址测试
#     ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
#     sin, sout, serr = ssh_huitie.exec_command('cd /mnt/ssd;echo $?', timeout=10, get_pty=True)
#     res = sout.read().decode('utf-8').strip()[-1]
#     if int(res) != 0:
#         print('>> 5G网口业务测试：创建ssd，挂载sda')
#         sin, sout, serr = \
#             ssh_huitie.exec_command('mkfs.ext4 -F /dev/sda;mkdir /mnt/ssd;mount /dev/sda /mnt/ssd;echo $?',
#                                     timeout=60, get_pty=True)
#         res = sout.read().decode('utf-8').strip()
#         print(res)
#     else:
#         sin, sout, serr = ssh_huitie.exec_command('umount /mnt/ssd;rm -rf /mnt/ssd')
#         res = sout.read().decode('utf-8').strip()
#         if not res:
#             print('>> 5G网口业务测试：取消挂载，删除ssd文件，准备重新cp文件')
#             print('>> 5G网口业务测试：创建ssd，挂载sda')
#             sin, sout, serr = \
#                 ssh_huitie.exec_command('mkfs.ext4 -F /dev/sda;mkdir /mnt/ssd;mount /dev/sda /mnt/ssd;echo $?',
#                                         timeout=60, get_pty=True)
#             res = sout.read().decode('utf-8').strip()
#             print(res)
#     ssh_huitie.close()
#
#     ssh_server.exec_command('lftp -u root,HTGW@123# -e "set file:charset utf-8;cd /mnt/ssd/;'
#                             'mirror -R -c --parallel=5 /home/dg-fpga/dateback/.;exit" 192.168.1.101:21',
#                             timeout=60, get_pty=True)
#
#     sar_cmd1 = "sar -n DEV 1 | grep eth1"
#     sar_cmd2 = "sar -n DEV 1 | grep eth2"
#     lftp_cmd1 = "lftp -u dg-fpga,fpga*1234 -e \"set file:charset utf-8;cd /update/;mirror -R -c --parallel=5 /mnt/ssd/.;exit\" 192.168.1.246:21"
#     lftp_cmd2 = "lftp -u dg-fpga,fpga*1234 -e \"set file:charset utf-8;cd /update/;mirror -R -c --parallel=5 /mnt/ssd/.;exit\" 192.168.2.246:21"
#     sar_cmds = [sar_cmd1, sar_cmd2]
#     lftp_cmds = [lftp_cmd1, lftp_cmd2]
#
#     for sar_cmd, lftp_cmd in zip(sar_cmds, lftp_cmds):
#         # ssh并发登录慧铁设备
#         ssh_huitie = ssh_connect(huitie_ip, huitie_user, huitie_password)
#         print(f'The sar cmd is {sar_cmd}')
#         print(f'The lftp cmd is {lftp_cmd}')
#         threads = []
#         thread1 = threading.Thread(target=lftp, args=(cli_2nd, lftp_cmd), name='lftp')
#         thread2 = threading.Thread(target=sar, args=(cli_2nd, sar_cmd), name='Sar')
#         threads.append(thread1)
#         threads.append(thread2)
#
#         for t in threads:
#             t.start()
#             print("Start {}".format(t.name))
#
#         for t in threads:
#             t.join()
#         print("文件传输完成，删除update下的文件")
#         restore_default(s_srv)
#         print('>> 退出所有线程')
#         time.sleep(5)
#
#     ssh_server.close()
#     ssh_huitie.close()
#
#     # 计算lftp速率上传平均值
#     data_list = get_data()
#     avg_rate = get_avg(data_list)
#     print('>> 文件传输速率：' + avg_rate)





script_file_path = '.\\'  # 脚本文件位置

huitie_ip = '192.168.1.101'
server_ip_list = ['192.168.1.246', '192.168.2.246', '192.168.3.246',
                  '192.168.4.246', '192.168.5.246', '192.168.6.246']
huitie_user = 'root'
huitie_password = 'HTGW@123#'
server_user = 'dg-fpga'
server_password = 'fpga*1234'

"""调试信息↓ ↓ ↓ ↓ ↓"""
# huitie_ip = '192.168.255.11'
# server_ip_list = ['192.168.255.31', '192.168.255.31', '192.168.255.31',
#                   '192.168.255.31', '192.168.255.31', '192.168.255.31']
# huitie_user = 'dg'
# huitie_password = 'passw0rd'
"""调试信息↑ ↑ ↑ ↑ ↑"""



"""log设置"""
logger = logging.getLogger('test')
logger.setLevel(logging.INFO)
def log_info():
    logger_use(script_file_path + 'log/')  # 日志存放位置


def excel_set():
    """excel设置"""
    global sn_num_input
    sn_num_input = open('.\\config', 'r', encoding='utf-8').readline()
    print('>> SN:', sn_num_input)
    global excel_report
    example_file_name = script_file_path + 'excel_example/huitie_report_example.xlsx'

    # 判断是否已存在测试报告
    global excel_report
    if os.path.exists(script_file_path + '/report/huitie_SN' + sn_num_input + '.xlsx'):
        excel_report = ExcelOpenPyXl(script_file_path + '/report/huitie_SN' + sn_num_input + '.xlsx')  # excel测试报告
        excel_report.get_sheet('Sheet1')
    else:
        excel_use = ExcelOpenPyXl(example_file_name)
        excel_file_name = time.strftime('huitie_SN' + sn_num_input + '.xlsx')
        excel_use.save_excel(script_file_path + 'report/' + excel_file_name)
        excel_use.close_excel()

        excel_report = ExcelOpenPyXl(script_file_path + 'report/' + excel_file_name)  # excel测试报告
        excel_report.get_sheet('Sheet1')


ssh = paramiko.SSHClient()  # 创建ssh对象
