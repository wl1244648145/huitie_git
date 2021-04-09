import os
import sys
from pathlib import Path
import threading
import time
import psutil
import re
current_folder = Path(__file__).absolute().parent
os.chdir(str(current_folder))
sys.path.append(str(current_folder.parent))
from utils import util

# 参数列表
ht_dev = '192.168.1.101'
srv_dev = [
    '192.168.1.246', '192.168.2.246', '192.168.3.246', '192.168.4.246',
    '192.168.5.246', '192.168.6.246'
]
user = 'root'
pwd = 'HTGW@123#'
user_srv = 'dg-fpga'
pwd_srv = 'fpga*1234'

# Step 6
pre_up = "lftp -u root,HTGW@123# -e \"set file:charset utf-8;cd /mnt/ssd/;mirror -R -c --parallel=5 /home/dg-fpga/dateback/.;exit\" 192.168.1.101:21"
sar_cmd1 = "sar -n DEV 1 | grep eth1"
sar_cmd2 = "sar -n DEV 1 | grep eth2"
lftp_cmd1 = "lftp -u dg-fpga,fpga*1234 -e \"set file:charset utf-8;cd /update/;mirror -R -c --parallel=5 /mnt/ssd/.;exit\" 192.168.1.246:21"
lftp_cmd2 = "lftp -u dg-fpga,fpga*1234 -e \"set file:charset utf-8;cd /update/;mirror -R -c --parallel=5 /mnt/ssd/.;exit\" 192.168.2.246:21"
sar_cmds = [sar_cmd1, sar_cmd2]
lftp_cmds = [lftp_cmd1, lftp_cmd2]

try:
    client = util.ssh_connect(ht_dev, user, pwd)
    s_srv = util.ssh_connect(srv_dev[0], user_srv, pwd_srv)
except:
    print('请确认ssh连接是否正常')

# print('\n======================开始进行第4章的测试=================================')


def test_ping_01():
    # ssh登录慧铁设备，ping远端各个server ip
    for srv in srv_dev:
        ping_cmd = 'ping -c 1 {}'.format(srv)
        sin, sout, serr = client.exec_command(ping_cmd,
                                              timeout=60,
                                              get_pty=True)
        res = sout.read().strip().decode()
        assert 'ttl' in res


def test_iperf_10g_02():
    # 登录远端server，开启监听模式
    ssh_server = util.ssh_connect(srv_dev[0], user_srv, pwd_srv)
    iperf_lis = 'iperf3 -s'
    ssh_server.exec_command(iperf_lis, timeout=60, get_pty=True)
    # client端通过iperf测试10G网口速率
    for srv in srv_dev[0:2]:
        iperf_speed = f'iperf3 -c {srv} -t 30'
        sin, sout, serr = client.exec_command(iperf_speed,
                                              timeout=60,
                                              get_pty=True)
        for line in util.line_buffered(sout):
            with open(r"E:\iperf.txt", 'w') as fd:
                fd.write(line)
        # 结果判定
        speed_sender = util.get_sender()
        speed_receiver = util.get_receiver()
        print(f"The speed in 10g is {speed_sender} and {speed_receiver}")
        assert float(speed_sender) > 3.0 and float(speed_receiver) > 3.0
    ssh_server.close()


def test_iperf_1g_03():
    # 登录远端server，开启监听模式
    ssh_server = util.ssh_connect(srv_dev[0], user_srv, pwd_srv)
    iperf_lis = 'iperf3 -s'
    ssh_server.exec_command(iperf_lis, timeout=60, get_pty=True)
    # client端通过iperf测试1G网口速率
    for srv in srv_dev[2:6]:
        iperf_speed = f'iperf3 -c {srv} -t 30'
        sin, sout, serr = client.exec_command(iperf_speed,
                                              timeout=60,
                                              get_pty=True)
        for line in util.line_buffered(sout):
            with open(r"E:\iperf.txt", 'w') as fd:
                fd.write(line)
        # 结果判定
        speed_sender = util.get_sender()
        speed_receiver = util.get_receiver()
        print(f"The speed in 1g is {speed_sender} and {speed_receiver}")
        assert int(speed_sender) > 900 and int(speed_receiver) > 900
    ssh_server.close()


# print('\n======================开始进行第5章的测试=================================')


def test_rtc_04():
    date_get = b'date'
    date_set = b'date -s "20210120 17:00:00"'
    date_save = b'hwclock -w'
    date_save_check = b'hwclock'

    # Get current time
    sin, sout, serr = client.exec_command(date_get)
    res = sout.read().strip().decode()
    assert 'UTC' in res

    # Set new timeout=60,
    client.exec_command(date_set)

    # Check new time config
    sin, sout, serr = client.exec_command(date_get)
    ret = sout.read().strip().decode()
    assert '17:00:00' in ret and '2021' in ret and 'Jan' in ret

    # Save time config
    client.exec_command(date_save)
    time.sleep(5)
    sin, sout, serr = client.exec_command(date_save_check,
                                          timeout=10,
                                          get_pty=True)
    r = sout.read().strip().decode()
    print(r)
    # while True:
    #     sin, sout, serr = client.exec_command(date_save_check, timeout=10, get_pty=True)
    #     r = sout.read().strip().decode()
    #     if 'Cannot' not in r and 'Invalid' not in r and 'error' not in r and 'Time out' not in r:
    #         break
    # print(r)
    assert '2021-01-20' in r


def test_temp_05():
    temp_get = b'cat /sys/class/hwmon/hwmon0/temp1_input'
    sin, sout, serr = client.exec_command(temp_get)
    r = sout.read().strip().decode()
    assert int(r) > 1000


def test_eeprom_06():
    htid_set = b'htgw id HTKJ000100'
    htid_get = b'htgw id'
    # 配置htgw idid
    sin, sout, serr = client.exec_command(htid_set, get_pty=True)
    res = sout.read().strip().decode()
    print(res)
    
    client.exec_command(b'reboot')
    time.sleep(20)
    while True:
        ret = util.check_connection(ht_dev)
        time.sleep(5)
        print('Wait!!!!!!!!!!!')
        if ret:
            break
    
    time.sleep(20)
    chient = util.ssh_connect(ht_dev, user, pwd)
    # 查看新配置htgw id是否生效
    sin, sout, serr = chient.exec_command(htid_get)
    r = sout.read().strip().decode()
    assert 'HTKJ000100' in r
    client.close()


def test_usbm_07():
    usb_check = b'dmesg -c'
    sda_check = b'fdisk -l'

    # Test sda
    print('\n=========Start check the sda on HT================')
    # global client_new
    client_new = util.ssh_connect(ht_dev, user, pwd)
    sin, sout, serr = client_new.exec_command(sda_check, get_pty=True)
    res = sout.read().strip().decode()
    print(type(res))
    print(res)
    assert 'Disk /dev/sda' in res

    # Test usb2.0
    print('\n=========Start test usb2.0================')
    while True:
        r = input('请插入USB2.0盘至慧铁的USB_M口，输入y或者yes继续测试: ')
        if 'yes' in r or 'y' in r:
            print('Continue the usb_m test!')
            break
    sin, sout, serr = client_new.exec_command(usb_check, get_pty=True)
    res = sout.read().strip().decode()
    assert 'new high-speed USB device' in res

    # Test usb3.0
    print('\n=========Start test usb3.0================')
    while True:
        r = input('请插入USB3.0盘至慧铁的USB_M口，输入y或者yes继续测试: ')
        if 'yes' in r or 'y' in r:
            print('Continue the usb_m test!')
            break
    sin, sout, serr = client_new.exec_command(usb_check, get_pty=True)
    res = sout.read().strip().decode()
    assert 'new SuperSpeed Gen 1 USB device' in res


def test_virtual_usb_08():
    client_new = util.ssh_connect(ht_dev, user, pwd)
    disk_nums_before = len(psutil.disk_partitions())
    cmds_config = [
        b'depmod', b'modprobe libcomposite', b'modprobe usb_f_mass_storage',
        b'modprobe g_mass_storage file=/dev/sda removable=1'
    ]
    for cmd in cmds_config:
        sin, sout, serr = client_new.exec_command(cmd, get_pty=True)
    while True:
        r = input('\n请将USB双公头延长线一段接入测试电脑，输入y或者yes继续测试: ')
        if 'yes' in r or 'y' in r:
            print('Continue the virtual_usb_m test!')
            break
    disk_nums_after = len(psutil.disk_partitions())
    assert disk_nums_after == disk_nums_before + 1


def test_mac_09():
    client_new = util.ssh_connect(ht_dev, user, pwd)
    def get_mac():
        """获取所有eth的MAC地址，并返回列表保存"""
        cmd_get_mac = b'ip link'
        sin, sout, serr = client_new.exec_command(cmd_get_mac, get_pty=True)
        res = sout.read().strip().decode()
        eth_ports = ['eth1', 'eth2', 'eth3', 'eth4', 'eth5', 'eth6']
        eth_macs = []
        for port in eth_ports:
            eth_mac = re.search(f'.*{port}:.*\n.*link/ether (.*) brd.*', res)
            if eth_mac:
                eth_macs.append(eth_mac.group(1))
            else:
                print('Failed to get MAC address')
        return eth_macs

    def get_id_in_mac(mac_group):
        """获取MAC地址48:54:00:00:7b:01中的00:00:7b"""
        id_mac = mac_group[0]
        id_in_mac = id_mac[6:14]
        return id_in_mac

    def get_htgwid():
        """通过htgw id获取HTJK0000100中的000100"""
        sin, sout, serr = client_new.exec_command(b'htgw id', get_pty=True)
        res = sout.read().strip().decode()
        htgwid_group = re.search(f'ID: HTKJ(.*).*', res)
        print(htgwid_group)
        htgwid = htgwid_group.group(1)
        return htgwid

    eth_mac_group = get_mac()
    id_in_mac = get_id_in_mac(eth_mac_group)
    print(f'The htgw id in mac is {id_in_mac}')

    # 判定所有的eth mac地址中存在相同的htgw id
    for mac in eth_mac_group:
        assert id_in_mac in mac

    # htgw id 获取ID
    htgwid_dec = get_htgwid()
    # 10进制id转换16进制
    htgwid_hex = hex(int(htgwid_dec))
    # 删除0x，保留剩余位数
    htgw_hex_id = htgwid_hex.replace('0x', '')
    # 将MAC中的‘：’删除，并重新组装MAC中的htgw id，16进制
    id_group = id_in_mac.split(':')
    id_hex = ''.join(id_group)
    # 判定htgw id获取的id与各端口的MAC中的htgw id一致
    assert htgw_hex_id in id_hex


# print('\n======================开始进行第6章的测试=================================')


def test_lftp_5g_100():
    """多线程实现sar和lftp"""
    # 确认ht设备/mnt/ssd下是否有待上传目录
    try:
        s_srv = util.ssh_connect(srv_dev[0], user_srv, pwd_srv)
    except:
        print('请确认ssh连接是否正常')
        s_srv = util.ssh_connect(srv_dev[0], user_srv, pwd_srv)
    client_new = util.ssh_connect(ht_dev, user, pwd)
    res = util.check_ssd(client_new)
    if res:
        print("Start to copy files to ssd")
        s_srv.exec_command(pre_up, timeout=60, get_pty=True)
        time.sleep(60)
        while True:
            sin, sout, serr = client_new.exec_command('cd /mnt/ssd;du -sh')
            r = sout.read().decode('utf').strip()
            if len(r) == 5:
                r = r[0:2]
            else:
                r = r[0:3]
            if float(r) < 15.0:
                time.sleep(10)
            else:
                break
        print('All done for pre-condition')
    for sar_cmd, lftp_cmd in zip(sar_cmds, lftp_cmds):
    # for sar_cmd, lftp_cmd in ((sar_cmd1, lftp_cmd1), (sar_cmd2, lftp_cmd2)):
        cli_2nd = util.ssh_connect(ht_dev, user, pwd)
        print(f'The sar cmd is {sar_cmd}')
        print(f'The lftp cmd is {lftp_cmd}')
        threads = []
        thread1 = threading.Thread(target=util.lftp,
                                   args=(cli_2nd, lftp_cmd),
                                   name='lftp')
        thread2 = threading.Thread(target=util.sar,
                                   args=(cli_2nd, sar_cmd),
                                   name='Sar')
        threads.append(thread1)
        threads.append(thread2)

        for t in threads:
            t.start()
            print("Start {}".format(t.name))

        for t in threads:
            t.join()
        print("lftp upload files complete, start to delete files on server!")
        util.restore_default(s_srv)
        print("Exit all the threads")
        time.sleep(5)

        #计算lftp速率上传平均值
        data_list = util.get_data()
        avg_rate = util.get_avg(data_list)
        print(avg_rate)
        assert float(avg_rate) > 1.65

    client_new.exec_command('umount /mnt/ssd;rm -rf /mnt/ssd')
