## 部署流程

1. 安装miniconda

	[安装包下载](https://repo.anaconda.com/miniconda/)

2. 创建环境并切换

	```shell
	conda create --name occ python=3.8
	conda activate occ
	```

3. 安装Python第三方库

	```shell
	conda install -c conda-forge pythonocc-core=7.5.1
	pip install -r requirements.txt
	```

	requirements.txt:

	```requirements.txt
	PyQt5==5.15.11
	cx_Freeze==7.2.0
	setuptools==65.6.3
	```

4. 打包

	`python setup.py build `

	执行成功后，可执行程序在build目录下

5. 修改config.ini为可执行程序路径。例如

	```ini
	[Paths]
	executablePath = ./exe.macosx-11.0-arm64-3.8/utils
	```

6. 修改CMakeLists.txt中Qt5路径并在build目录下执行编译链接命令：

	```shell
	cd build
	cmake ..
	make
	```

	