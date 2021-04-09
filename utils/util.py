import paramiko
import subprocess
import sys
import logging
import re
import time
import os
from datetime import datetime

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)


def check_connection(host):
    ret = None
    cmd = f'ping -n 1 -w 500 -l 1 {host}'
    response = subprocess.Popen(cmd,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT,
                                stdin=subprocess.DEVNULL,
                                shell=True).communicate(timeout=30)[0].decode(
                                    sys.getdefaultencoding(),
                                    errors='replace').encode().decode(
                                        'utf-8', errors='replace')
    if 'ttl=' in response.lower():
        logger.debug("ping succeeded for {0}".format(host))
        ret = True
    else:
        logger.info("ping failed: cannot connect to device {0}".format(host))
        ret = False
    return ret


def ssh_connect(host, user, passwd):
    """ssh登录设备"""
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname=host, username=user, password=passwd, timeout=60)
    return ssh


def check_ssd(ssh):
    # sin, sout, serr = ssh.exec_command('rm -rf /mnt/ssd;mkdir /mnt/ssd;mount /dev/sda /mnt/ssd;echo $?',
                                       # timeout=10,
                                       # get_pty=True)
    # res = sout.read().decode('utf-8').strip()[-1]
    sin, sout, serr = ssh.exec_command('cd /mnt/ssd;echo $?',
                                       timeout=10,
                                       get_pty=True)
    res = sout.read().decode('utf-8').strip()[-1]
    if int(res) != 0:
        print('Start to mnount ssd')
        sin, sout, serr = ssh.exec_command(
            'mkfs.ext4 -F /dev/sda;mkdir /mnt/ssd;mount /dev/sda /mnt/ssd;echo $?',
            timeout=60,
            get_pty=True)
        ret = sout.read().decode('utf-8').strip()
        print(ret)
    else:
        sin, sout, serr = ssh.exec_command('ls /mnt/ssd')
        r = sout.read().decode('utf-8').strip()
        if not r:
            print("No files under ssd, please cp files for preparation!")
    return True


def line_buffered(f):
    """打印log解析"""
    line_buf = ""
    while not f.channel.exit_status_ready():
        line_buf += f.read(1024).decode()
        if line_buf.endswith('\n'):
            yield line_buf
            line_buf = ''


def get_sender():
    iperf_sender = None
    with open(r"E:\iperf.txt") as fd:
        lines = fd.readlines()
        for line in lines:
            if '[  4]   0.00-30.00' in line and 'sender' in line:
                sum_res = line.strip()
                match_sender = re.search(
                    '\[  4\].*GBytes\s+(\d+)\sMbits.*sender', sum_res)
                if match_sender:
                    iperf_sender = match_sender.group(1)
                else:
                    iperf_sender = re.search(
                        '\[  4\].*GBytes\s+(.*)\sGbits.*sender',
                        sum_res).group(1)
    return iperf_sender


def get_receiver():
    iperf_sender = None
    with open(r"E:\iperf.txt") as fd:
        lines = fd.readlines()
        for line in lines:
            if '[  4]   0.00-30.00' in line and 'receiver' in line:
                sum_res = line.strip()
                match_sender = re.search(
                    '\[  4\].*GBytes\s+(\d+)\sMbits.*receiver', sum_res)
                if match_sender:
                    iperf_sender = match_sender.group(1)
                else:
                    iperf_sender = re.search(
                        '\[  4\].*GBytes\s+(.*)\sGbits.*receiver',
                        sum_res).group(1)
    return iperf_sender


def sar(ssh, sar_cmd):
    """设备上通过sar命令读取网口速率"""
    sin, sout, serr = ssh.exec_command(sar_cmd, timeout=60, get_pty=True)
    dt = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_dir_format = str(datetime.now().month) + str(datetime.now().day) + str(
        datetime.now().year)
    log_name = f'lftp_{dt}.log'
    for line in line_buffered(sout):
        if not os.path.exists(f"E:\\lftp_logs_{log_dir_format}"):
            os.makedirs(f"E:\\lftp_logs_{log_dir_format}")
        with open(f"E:\\lftp_logs_{log_dir_format}\\{log_name}", 'a+') as fd:
            fd.write(line)
    # ssh.close()


def lftp(ssh, lftp_cmd):
    """通过lftp命令上传大文件至指定目的地"""
    timeout = 250
    endtime = time.time() + timeout
    sin, sout, serr = ssh.exec_command(lftp_cmd, timeout=60, get_pty=True)
    while not sout.channel.eof_received:
        time.sleep(1)
        if time.time() > endtime:
            sout.channel.close()
            break
    # res = sout.read().decode()
    # if len(res) == 0:
    ssh.close()


def restore_default(ssh_client):
    """文件传输结束后，目标主机上将大文件删除，以备下次传输测试"""
    rm_cmd = r'cd ~/update;rm -rf *'
    ssh_client.exec_command(rm_cmd, timeout=60)
    print("Complete deleting files on server!")


def get_data():
    data_l = []
    log_format = str(datetime.now().month) + str(datetime.now().day) + str(
        datetime.now().year)
    log_dir = f'E:\\lftp_logs_{log_format}'
    print(log_dir)
    for dirpath, dirnames, filenames in os.walk(log_dir):
        target_file = log_dir + '\\' + filenames[-1]
        with open(target_file, 'r') as fd:
            for line in fd:
                line_l = line.strip().split()
                if (len(line_l)) != 0:
                    data_l.append(line_l[5])
                else:
                    continue
    return data_l


def get_avg(p_list):
    sum_speed = 0
    for i in p_list:
        sum_speed = sum_speed + float(i)
    avg = sum_speed * 8 / len(p_list) / 1024 / 1024
    return avg
