# build_nuitka.py
import os
import subprocess
import sys

def run_command(command):
    print(f"执行命令: {command}")
    process = subprocess.Popen(command, shell=True)
    process.wait()
    if process.returncode != 0:
        print(f"命令执行失败，返回码: {process.returncode}")
        sys.exit(1)

# visualize_xlsx.py
print("正在打包visualize_xlsx.py...")
run_command("nuitka --standalone --show-progress --show-memory visualize_xlsx.py")

# 打包visualize_xlsx.py
print("正在打包visualize_xlsx.py...")
run_command("nuitka --standalone --show-progress --show-memory visualize_xlsx.py")

print("所有打包任务完成！")