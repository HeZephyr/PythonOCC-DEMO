import sys
import argparse
import xml.etree.ElementTree as ET
import os
import logging
import traceback
from datetime import datetime
from OCC.Core.Quantity import Quantity_Color
from OCC.Core._Quantity import Quantity_TOC_RGB
from PyQt5.QtCore import QT_VERSION_STR
from OCC.Display.backend import load_backend
from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_Reader, STEPControl_AsIs
from OCC.Core.IGESControl import IGESControl_Reader, IGESControl_Writer
from OCC.Core.Interface import Interface_Static_SetCVal
from OCC.Core.IFSelect import IFSelect_RetDone

from OCC.Core.V3d import V3d_Zneg, V3d_Yneg, V3d_Xneg

load_backend("qt-pyqt5")
from OCC.Display.qtDisplay import qtViewer3d

from PyQt5.QtWidgets import (
    QApplication, QTreeWidget, QTreeWidgetItem, QWidget, QMainWindow,
    QHBoxLayout, QVBoxLayout, QDesktopWidget, QPushButton, QFileDialog,
    QLabel, QComboBox, QGroupBox, QMessageBox, QProgressDialog,
    QTextEdit, QStatusBar
)
from PyQt5.QtCore import Qt, QTimer, QCoreApplication
from PyQt5.QtGui import QFont
from OCC.Core.BRepPrimAPI import BRepPrimAPI_MakeCylinder, BRepPrimAPI_MakeSphere
from OCC.Core.gp import gp_Pnt, gp_Dir, gp_Ax2, gp_Vec
from OCC.Core.TopoDS import TopoDS_Shape, TopoDS_Compound
from OCC.Core.BRep import BRep_Builder
from OCC.Core.AIS import AIS_Shape, AIS_InteractiveContext
from OCC.Core.Quantity import Quantity_NOC_BLUE, Quantity_NOC_YELLOW, Quantity_NOC_RED, Quantity_NOC_GREEN


def setup_logging():
    """设置日志记录"""
    log_dir = "logs"
    try:
        os.makedirs(log_dir, exist_ok=True)
        if not os.access(log_dir, os.W_OK):
            raise PermissionError(f"无写入权限: {log_dir}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"xml_app_{timestamp}.log")

        # 统一日志格式
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # 文件处理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.DEBUG)

        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)

        # 配置根日志器
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.DEBUG)
        root_logger.addHandler(file_handler)
        root_logger.addHandler(console_handler)

        # 创建应用专属日志器
        app_logger = logging.getLogger("visualize_xml")
        app_logger.propagate = False
        app_logger.setLevel(logging.DEBUG)
        app_logger.addHandler(file_handler)
        app_logger.addHandler(console_handler)

        return log_file
    except Exception as e:
        print(f"致命错误: 无法初始化日志系统 - {str(e)}")
        traceback.print_exc()
        sys.exit(1)

# 全局日志器
logger = logging.getLogger("visualize_xml")


class MainWindow(QWidget):
    def __init__(self, xml_file=None):
        super().__init__()
        self.setWindowTitle("航电布线可视化系统")
        
        logger.info("初始化应用程序...")
        
        # 初始化消息框和计时器
        self.msg_box = None
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer_connected = False
        
        # 高亮和颜色管理
        self.ais_shapes = {}  # 存储所有AIS对象，用于颜色管理
        self.highlighted_shapes = []  # 当前高亮的形状IDs
        
        # 节点相关数据结构
        self.unique_nodes = {}  # 存储唯一节点信息，使用name作为键，(x, y, z)作为值
        self.node_shapes = []  # 存储节点的 TopoDS_Shape 对象
        self.node_id_map = {}  # Map node_name to node_index ('node_0', 'node_1', etc.)
        
        # 存储链接相关的数据，用于查询
        self.link_data = {}  # 存储链接数据，使用索引作为键
        self.node_to_links = {}  # 存储节点关联的链接，使用节点name作为键
        self.shape_to_info = {}  # 存储形状ID到信息的映射 {shape_id: info_dict}
        
        # 设置字体
        font = QFont()
        font.setPointSize(10)
        self.setFont(font)

        # 主布局
        main_layout = QVBoxLayout()  # 使用垂直布局来添加状态栏

        # 水平布局用于左侧面板和3D视图
        horizontal_layout = QHBoxLayout()

        # 左侧布局
        left_layout = QVBoxLayout()

        # 树状结构
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["模型结构"])
        self.tree.setColumnWidth(0, 300)
        left_layout.addWidget(self.tree)

        # 文件操作分组
        file_group = QGroupBox("文件操作")
        file_layout = QVBoxLayout()
        
        # 导入部分
        import_layout = QHBoxLayout()
        import_layout.addWidget(QLabel("导入文件:"))
        self.file_format_combo = QComboBox()
        self.file_format_combo.addItems(["STEP", "IGES"])
        import_layout.addWidget(self.file_format_combo)
        self.import_button = QPushButton("导入文件")
        self.import_button.clicked.connect(self.import_file)
        import_layout.addWidget(self.import_button)
        file_layout.addLayout(import_layout)

        # 导出部分
        export_layout = QHBoxLayout()
        export_layout.addWidget(QLabel("导出文件:"))
        self.export_format_combo = QComboBox()
        self.export_format_combo.addItems(["STEP", "IGES"])
        export_layout.addWidget(self.export_format_combo)
        self.export_button = QPushButton("导出文件")
        self.export_button.clicked.connect(self.export_file)
        export_layout.addWidget(self.export_button)
        file_layout.addLayout(export_layout)
        
        file_group.setLayout(file_layout)
        left_layout.addWidget(file_group)
        
        # 布局操作分组
        layout_group = QGroupBox("布局操作")
        layout_layout = QVBoxLayout()
        # 居中按钮
        self.layout_button = QPushButton("模型居中")
        layout_layout.addWidget(self.layout_button)
        # 正视图按钮
        self.front_view_button = QPushButton("正视图")
        layout_layout.addWidget(self.front_view_button)
        # 俯视图按钮
        self.top_view_button = QPushButton("俯视图")
        layout_layout.addWidget(self.top_view_button)
        # 右视图按钮
        self.right_view_button = QPushButton("右视图")
        layout_layout.addWidget(self.right_view_button)
        
        layout_group.setLayout(layout_layout)
        left_layout.addWidget(layout_group)

        # 添加信息显示区域
        info_group = QGroupBox("对象信息")
        info_layout = QVBoxLayout()
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setMinimumHeight(100)
        info_layout.addWidget(self.info_text)
        info_group.setLayout(info_layout)
        left_layout.addWidget(info_group)

        horizontal_layout.addLayout(left_layout, 1)

        # 3D 视图
        self.viewer = qtViewer3d(self)
        bg_color = Quantity_Color(0.8, 0.8, 0.8, Quantity_TOC_RGB)
        # Access the underlying AIS_InteractiveContext
        self.context: AIS_InteractiveContext = self.viewer._display.Context
        self.viewer._display.View.SetBackgroundColor(bg_color)
        horizontal_layout.addWidget(self.viewer, 2)
        
        # 添加主水平布局
        main_layout.addLayout(horizontal_layout)

        # 添加状态栏
        self.status_bar = QStatusBar()
        self.status_bar.setFixedHeight(25)
        main_layout.addWidget(self.status_bar)

        self.setLayout(main_layout)
        
        # 布局操作绑定
        self.layout_button.clicked.connect(self.viewer._display.FitAll)
        # 正视图
        self.front_view_button.clicked.connect(self.set_front_view)
        # 俯视图
        self.top_view_button.clicked.connect(self.set_top_view)
        # 右视图
        self.right_view_button.clicked.connect(self.set_right_view)
        # 存储线段形状和原始颜色
        self.segment_shapes = []
        self.segments = []
        
        # 保存导入的STEP/IGES模型
        self.step_shapes = {}  # 形状ID到形状对象的映射
        self.main_shape = None
        
        # 树项点击事件
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        self.selected_item = None
        
        # 设置交互功能
        self.setup_interaction()

        # 如果有XML文件，则解析并显示
        if xml_file is not None:
            try:
                self.parse_xml_and_populate_tree(xml_file)
                # 在布局完成后绘制模型
                QTimer.singleShot(100, self.draw_segments)  # 延迟100毫秒确保UI完全初始化
                logger.info("数据加载和初始化完成")
            except Exception as e:
                logger.error(f"初始化数据时出错: {str(e)}")
                logger.error(traceback.format_exc())
                QMessageBox.critical(self, "初始化错误", f"加载数据时出现错误: {str(e)}")

    # 定义设置正视图的函数
    def set_front_view(self):
        self.viewer._display.View.SetProj(V3d_Zneg)
        self.viewer._display.FitAll()

    # 定义设置俯视图的函数
    def set_top_view(self):
        self.viewer._display.View.SetProj(V3d_Yneg)
        self.viewer._display.FitAll()

    # 定义设置右视图的函数
    def set_right_view(self):
        self.viewer._display.View.SetProj(V3d_Xneg)
        self.viewer._display.FitAll()

    def parse_xml_and_populate_tree(self, file_path):
        """解析XML文件并构建树结构，支持多种XML格式"""
        try:
            self.xml_file_path = file_path
            logger.info(f"正在解析XML文件: {file_path}")
            self.status_bar.showMessage(f"正在解析XML文件: {file_path}")
            QApplication.processEvents()
            
            # 解析XML文件
            try:
                tree = ET.parse(file_path)
                self.root = tree.getroot()
            except Exception as e:
                logger.error(f"XML解析错误: {str(e)}")
                QMessageBox.critical(self, "解析错误", f"解析XML文件时出错: {str(e)}")
                return
            
            # 清空之前的数据
            self.tree.clear()
            self.segments.clear()
            self.unique_nodes.clear()
            self.link_data.clear()
            self.node_to_links.clear()
            self.segment_shapes.clear()
            self.viewer._display.EraseAll()
            self.shape_to_info.clear()
            self.node_shapes.clear()
            self.node_id_map.clear()
            
            self.status_bar.showMessage("正在构建树形结构...")
            QApplication.processEvents()
            
            # 检测XML格式（根据根节点名称）
            root_tag = self.root.tag
            logger.info(f"检测到XML格式: {root_tag}")
            
            # 创建根节点
            root_item = QTreeWidgetItem(self.tree, [f"{root_tag}: {os.path.basename(file_path)}"])
            root_item.setExpanded(True)
            
            # 根据不同格式解析
            if root_tag == "MultiDeviceNet":
                self.parse_multi_device_net(root_item)
            elif root_tag == "TwoDeviceNet":
                self.parse_two_device_net(root_item)
            else:
                logger.warning(f"未知的XML格式: {root_tag}")
                QMessageBox.warning(self, "格式警告", f"未知的XML格式: {root_tag}，将尝试通用解析")
                # 回退到通用解析
                self.parse_generic_format(root_item)
            
            # 创建节点的3D形状
            self.create_node_shapes()
            
            self.status_bar.showMessage(f"文件 {os.path.basename(file_path)} 加载完成")
            logger.info(f"XML文件解析完成，包含 {len(self.unique_nodes)} 个节点和 {len(self.segments)} 条连接")
        
        except Exception as e:
            logger.error(f"解析XML和构建树时发生错误: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "处理错误", f"处理XML文件时出错: {str(e)}")

    def parse_multi_device_net(self, root_item):
        """解析MultiDeviceNet格式的XML"""
        logger.info("解析MultiDeviceNet格式")
        
        for net in self.root.findall("./Net"):
            net_name = net.get("name", "未命名网络")
            net_item = QTreeWidgetItem(root_item, [f"Net: {net_name}"])
            net_item.setData(0, Qt.UserRole, {"type": "net", "name": net_name})
            
            # 处理设备信息
            self.parse_devices(net, net_item)
            
            # 处理等电位点
            self.parse_isoelectric_points(net, net_item)
            
            # 检查是否有TotalNetwork
            total_network = net.find("TotalNetwork")
            if total_network is not None:
                logger.info(f"发现TotalNetwork in Net: {net_name}")
                self.parse_total_network(total_network, net_item)
            else:
                logger.info(f"未发现TotalNetwork in Net: {net_name}，解析SubNet")
                self.parse_subnets(net, net_item)

    def parse_two_device_net(self, root_item):
        """解析TwoDeviceNet格式的XML"""
        logger.info("解析TwoDeviceNet格式")
        
        for net in self.root.findall("./Net"):
            net_name = net.get("name", "未命名网络")
            net_item = QTreeWidgetItem(root_item, [f"Net: {net_name}"])
            net_item.setData(0, Qt.UserRole, {"type": "net", "name": net_name})
            
            # 检查是否有TotalNetwork
            total_network = net.find("TotalNetwork")
            if total_network is not None:
                logger.info(f"发现TotalNetwork in Net: {net_name}")
                self.parse_total_network(total_network, net_item)
            else:
                logger.info(f"未发现TotalNetwork in Net: {net_name}，解析SubNet")
                self.parse_subnets(net, net_item)

    def parse_generic_format(self, root_item):
        """通用格式解析（回退方案）"""
        logger.info("使用通用格式解析")
        
        for net in self.root.findall(".//Net"):
            net_name = net.get("name", "未命名网络")
            net_item = QTreeWidgetItem(root_item, [f"Net: {net_name}"])
            net_item.setData(0, Qt.UserRole, {"type": "net", "name": net_name})
            
            # 尝试解析各种可能的结构
            self.parse_devices(net, net_item)
            self.parse_isoelectric_points(net, net_item)
            
            total_network = net.find("TotalNetwork")
            if total_network is not None:
                self.parse_total_network(total_network, net_item)
            else:
                self.parse_subnets(net, net_item)

    def parse_devices(self, net, net_item):
        """解析设备信息"""
        devices = net.find("Devices")
        if devices is not None:
            devices_item = QTreeWidgetItem(net_item, ["设备"])
            devices_item.setData(0, Qt.UserRole, {"type": "devices"})
            
            for device in devices.findall("Device"):
                device_name = device.get("name", "未命名设备")
                x = float(device.get("X", 0))
                y = float(device.get("Y", 0))
                z = float(device.get("Z", 0))
                
                device_item = QTreeWidgetItem(devices_item, [f"设备: {device_name}"])
                device_item.setData(0, Qt.UserRole, {
                    "type": "device",
                    "name": device_name,
                    "x": x,
                    "y": y,
                    "z": z
                })
                
                # 添加设备到唯一节点列表
                self.unique_nodes[device_name] = (x, y, z)
                logger.debug(f"添加设备节点: {device_name} at ({x}, {y}, {z})")

    def parse_isoelectric_points(self, net, net_item):
        """解析等电位点信息"""
        iso_points = net.find("IsoelectricPoints")
        if iso_points is not None:
            isoe_item = QTreeWidgetItem(net_item, ["等电位点"])
            isoe_item.setData(0, Qt.UserRole, {"type": "isoe"})
            
            for iso_point in iso_points.findall("IsoelePt"):
                point_name = iso_point.get("name", "未命名等电位点")
                x = float(iso_point.get("X", 0))
                y = float(iso_point.get("Y", 0))
                z = float(iso_point.get("Z", 0))
                
                point_item = QTreeWidgetItem(isoe_item, [f"等电位点: {point_name}"])
                point_item.setData(0, Qt.UserRole, {
                    "type": "isopt",
                    "name": point_name,
                    "x": x,
                    "y": y,
                    "z": z
                })
                
                # 添加等电位点到唯一节点列表
                self.unique_nodes[point_name] = (x, y, z)
                logger.debug(f"添加等电位点节点: {point_name} at ({x}, {y}, {z})")

    def parse_total_network(self, total_network, net_item):
        """解析TotalNetwork节点"""
        total_network_item = QTreeWidgetItem(net_item, ["TotalNetwork"])
        total_network_item.setData(0, Qt.UserRole, {
            "type": "total_network",
            "name": "TotalNetwork"
        })
        
        # 存储所有TotalNetwork下的链接形状ID
        self.total_network_shapes = []
        segment_idx = len(self.segments)  # 当前段的起始索引
        
        # 处理TotalNetwork下的Network节点
        for network in total_network.findall("Network"):
            network_name = network.get("name", f"未命名网络{segment_idx}")
            network_item = QTreeWidgetItem(total_network_item, [f"Network: {network_name}"])
            network_item.setData(0, Qt.UserRole, {
                "type": "network",
                "name": network_name,
                "index": segment_idx
            })
            
            # 处理起点和终点
            start_point = network.find("StartPoint")
            end_point = network.find("EndPoint")
            
            if start_point is not None and end_point is not None:
                start_name = start_point.get("name", "未命名起点")
                start_x = float(start_point.get("x", 0))
                start_y = float(start_point.get("y", 0))
                start_z = float(start_point.get("z", 0))
                
                end_name = end_point.get("name", "未命名终点")
                end_x = float(end_point.get("x", 0))
                end_y = float(end_point.get("y", 0))
                end_z = float(end_point.get("z", 0))
                
                # 添加Start点和End点到节点列表
                self.unique_nodes[start_name] = (start_x, start_y, start_z)
                self.unique_nodes[end_name] = (end_x, end_y, end_z)
                
                # 添加连接信息
                link_info = {
                    "type": "link",
                    "start_node": start_name,
                    "end_node": end_name,
                    "start_pos": (start_x, start_y, start_z),
                    "end_pos": (end_x, end_y, end_z),
                    "network_name": network_name,
                    "parent": "TotalNetwork"
                }
                
                self.segments.append((
                    (start_x, start_y, start_z),
                    (end_x, end_y, end_z),
                    segment_idx
                ))
                
                # 存储链接数据
                self.link_data[segment_idx] = link_info
                
                # 记录节点相关的链接
                if start_name not in self.node_to_links:
                    self.node_to_links[start_name] = []
                self.node_to_links[start_name].append(segment_idx)
                
                if end_name not in self.node_to_links:
                    self.node_to_links[end_name] = []
                self.node_to_links[end_name].append(segment_idx)
                
                logger.debug(f"添加TotalNetwork链接: {network_name} ({start_name} -> {end_name})")
                segment_idx += 1
            else:
                logger.warning(f"Network {network_name} 缺少起点或终点")
                error_item = QTreeWidgetItem(network_item, ["错误: 缺少起点或终点"])

    def parse_subnets(self, net, net_item):
        """解析SubNet节点（无TotalNetwork情况）"""
        logger.info("解析SubNet结构")
        segment_idx = len(self.segments)  # 当前段的起始索引
        
        # 首先解析FromDeviceOrConnector和ToDeviceOrConnector（TwoDeviceNet格式）
        self.parse_device_connectors(net, net_item)
        
        for subnet in net.findall("SubNet"):
            subnet_name = subnet.get("name", "未命名子网")
            subnet_item = QTreeWidgetItem(net_item, [f"SubNet: {subnet_name}"])
            subnet_item.setData(0, Qt.UserRole, {"type": "subnet", "name": subnet_name})
            
            # 解析FromDeviceOrConnector和ToDeviceOrConnector（如果在SubNet级别）
            self.parse_device_connectors(subnet, subnet_item)
            
            # 解析Segement
            for segement in subnet.findall("Segement"):
                segement_name = segement.get("name", "未命名段")
                segement_item = QTreeWidgetItem(subnet_item, [f"Segement: {segement_name}"])
                segement_item.setData(0, Qt.UserRole, {"type": "segement", "name": segement_name})
                
                # 解析NetStartPoint和NetEndPoint
                net_start = segement.find("NetStartPoint")
                net_end = segement.find("NetEndPoint")
                
                if net_start is not None and net_end is not None:
                    start_device = net_start.get("name", "未知起始设备")
                    end_device = net_end.get("name", "未知终止设备")
                    
                    info_item = QTreeWidgetItem(segement_item, [f"路径: {start_device} -> {end_device}"])
                    info_item.setData(0, Qt.UserRole, {
                        "type": "route_info",
                        "start": start_device,
                        "end": end_device
                    })
                
                # 解析Network节点
                for network in segement.findall("Network"):
                    network_name = network.get("name", f"未命名网络{segment_idx}")
                    network_item = QTreeWidgetItem(segement_item, [f"Network: {network_name}"])
                    network_item.setData(0, Qt.UserRole, {
                        "type": "network",
                        "name": network_name,
                        "index": segment_idx
                    })
                    
                    # 处理起点和终点
                    start_point = network.find("StartPoint")
                    end_point = network.find("EndPoint")
                    
                    if start_point is not None and end_point is not None:
                        start_name = start_point.get("name", "未命名起点")
                        start_x = float(start_point.get("x", 0))
                        start_y = float(start_point.get("y", 0))
                        start_z = float(start_point.get("z", 0))
                        
                        end_name = end_point.get("name", "未命名终点")
                        end_x = float(end_point.get("x", 0))
                        end_y = float(end_point.get("y", 0))
                        end_z = float(end_point.get("z", 0))
                        
                        # 添加Start点和End点到节点列表
                        self.unique_nodes[start_name] = (start_x, start_y, start_z)
                        self.unique_nodes[end_name] = (end_x, end_y, end_z)
                        
                        # 添加连接信息
                        link_info = {
                            "type": "link",
                            "start_node": start_name,
                            "end_node": end_name,
                            "start_pos": (start_x, start_y, start_z),
                            "end_pos": (end_x, end_y, end_z),
                            "network_name": network_name,
                            "parent": f"SubNet:{subnet_name}",
                            "segement": segement_name
                        }
                        
                        self.segments.append((
                            (start_x, start_y, start_z),
                            (end_x, end_y, end_z),
                            segment_idx
                        ))
                        
                        # 存储链接数据
                        self.link_data[segment_idx] = link_info
                        
                        # 记录节点相关的链接
                        if start_name not in self.node_to_links:
                            self.node_to_links[start_name] = []
                        self.node_to_links[start_name].append(segment_idx)
                        
                        if end_name not in self.node_to_links:
                            self.node_to_links[end_name] = []
                        self.node_to_links[end_name].append(segment_idx)
                        
                        logger.debug(f"添加SubNet链接: {network_name} ({start_name} -> {end_name})")
                        segment_idx += 1
                    else:
                        logger.warning(f"Network {network_name} in SubNet 缺少起点或终点")
                        error_item = QTreeWidgetItem(network_item, ["错误: 缺少起点或终点"])

    def parse_device_connectors(self, parent_element, parent_item):
        """解析FromDeviceOrConnector和ToDeviceOrConnector节点"""
        from_device = parent_element.find("FromDeviceOrConnector")
        to_device = parent_element.find("ToDeviceOrConnector")
        
        if from_device is not None or to_device is not None:
            connectors_item = QTreeWidgetItem(parent_item, ["设备连接器"])
            connectors_item.setData(0, Qt.UserRole, {"type": "connectors"})
            
            if from_device is not None:
                device_name = from_device.get("name", "未命名起始设备")
                x = float(from_device.get("x", 0))
                y = float(from_device.get("y", 0))
                z = float(from_device.get("z", 0))
                
                from_item = QTreeWidgetItem(connectors_item, [f"起始设备: {device_name}"])
                from_item.setData(0, Qt.UserRole, {
                    "type": "from_device",
                    "name": device_name,
                    "x": x,
                    "y": y,
                    "z": z
                })
                
                # 添加到唯一节点列表
                self.unique_nodes[device_name] = (x, y, z)
                logger.debug(f"添加起始设备节点: {device_name} at ({x}, {y}, {z})")
            
            if to_device is not None:
                device_name = to_device.get("name", "未命名终止设备")
                x = float(to_device.get("x", 0))
                y = float(to_device.get("y", 0))
                z = float(to_device.get("z", 0))
                
                to_item = QTreeWidgetItem(connectors_item, [f"终止设备: {device_name}"])
                to_item.setData(0, Qt.UserRole, {
                    "type": "to_device",
                    "name": device_name,
                    "x": x,
                    "y": y,
                    "z": z
                })
                
                # 添加到唯一节点列表
                self.unique_nodes[device_name] = (x, y, z)
                logger.debug(f"添加终止设备节点: {device_name} at ({x}, {y}, {z})")

    def create_node_shapes(self):
        """创建代表节点的球体形状"""
        logger.info(f"创建节点形状，节点数量: {len(self.unique_nodes)}")
        self.node_shapes = []  # 重置节点形状列表
        self.node_id_map = {}  # 重置节点ID映射

        # 设置球体的半径
        radius = 25.0  # 球体的半径，可以根据需要调整

        for i, (node_name, node_pos) in enumerate(self.unique_nodes.items()):
            try:
                # 创建球体
                center = gp_Pnt(*node_pos)
                sphere = BRepPrimAPI_MakeSphere(center, radius).Shape()
                if sphere.IsNull():
                    logger.warning(f"创建节点 {node_name} 的球体失败")
                    continue
                
                # 添加到节点形状列表
                self.node_shapes.append(sphere)
                
                # 创建节点的唯一ID格式
                node_shape_id = f"node_{i}"
                
                # 存储节点名称到ID的映射
                self.node_id_map[node_name] = node_shape_id
                
                # 存储节点信息
                self.shape_to_info[node_shape_id] = {
                    "type": "node",
                    "name": node_name,
                    "coordinates": node_pos,
                    "connected_links": self.node_to_links.get(node_name, [])
                }
                
            except Exception as e:
                logger.error(f"创建节点 {node_name} 的形状时出错: {str(e)}")
                logger.error(traceback.format_exc())
                continue

        logger.info(f"节点 TopoDS_Shape 创建完成，成功创建: {len(self.node_shapes)}")

    def draw_segments(self):
        """绘制所有线段"""
        try:
            if not self.segments:
                logger.warning("没有线段可以绘制")
                self.status_bar.showMessage("没有线段可以绘制")
                return

            logger.info(f"正在绘制 {len(self.segments)} 条线段")
            self.status_bar.showMessage(f"正在绘制 {len(self.segments)} 条线段...")
            QApplication.processEvents()

            # 清除之前的线段
            for shape in self.segment_shapes:
                if self.viewer._display.IsDisplayed(shape):
                    self.viewer._display.Erase(shape)

            self.segment_shapes.clear()
            self.total_network_shapes = []  # 清除TotalNetwork相关的形状

            # 为进度显示预处理
            progress = QProgressDialog("绘制线段...", None, 0, len(self.segments), self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(1000)  # 1秒后显示
            
            # 绘制所有线段
            for i, (start, end, idx) in enumerate(self.segments):
                if progress.wasCanceled():
                    break
                    
                progress.setValue(i)
                QApplication.processEvents()
                
                try:
                    # 创建起点和终点
                    p1 = gp_Pnt(start[0], start[1], start[2])
                    p2 = gp_Pnt(end[0], end[1], end[2])
                    
                    # 创建方向和轴
                    v = gp_Vec(p1, p2)
                    direction = gp_Dir(v)
                    
                    # 计算长度和半径
                    length = v.Magnitude()
                    radius = 5.0  # 默认半径
                    
                    # 检查长度，防止太短
                    if length < 1e-5:
                        logger.warning(f"线段 {idx} 长度为 {length}，太短无法绘制")
                        continue
                    
                    # 创建圆柱体作为线段
                    axis = gp_Ax2(p1, direction)
                    cylinder = BRepPrimAPI_MakeCylinder(axis, radius, length).Shape()
                    
                    # 创建AIS对象
                    ais_shape = AIS_Shape(cylinder)
                    ais_shape.SetColor(Quantity_Color(0, 0, 1, Quantity_TOC_RGB))  # 蓝色
                    
                    # 将形状添加到Interactive Context
                    self.context.Display(ais_shape, False)
                    
                    # 直接使用索引作为形状ID（而不是尝试获取Handle）
                    shape_id = idx  # 使用线段的索引作为ID
                    self.segment_shapes.append(shape_id)
                    # 存储AIS对象，用于后续访问
                    self.ais_shapes[shape_id] = ais_shape
                    self.shape_to_info[shape_id] = {
                        "type": "link",
                        "index": idx,
                        **self.link_data.get(idx, {})
                    }
                    
                    # 如果这个线段是TotalNetwork的子网络，将其添加到total_network_shapes
                    link_info = self.link_data.get(idx, {})
                    if link_info.get("parent") == "TotalNetwork":
                        self.total_network_shapes.append(shape_id)
                        
                except Exception as e:
                    logger.error(f"绘制线段 {idx} 时出错: {str(e)}")
                    logger.error(traceback.format_exc())
                    
            progress.setValue(len(self.segments))
            
            # 绘制节点（如果有）
            if self.node_shapes:
                logger.info("开始绘制节点...")
                
                for i, sphere in enumerate(self.node_shapes):
                    try:
                        # 获取节点ID
                        node_shape_id = f"node_{i}"
                        
                        # 创建AIS形状用于显示
                        ais_sphere = AIS_Shape(sphere)
                        ais_sphere.SetColor(Quantity_Color(Quantity_NOC_RED))  # 设置为红色
                        
                        # 显示球体
                        self.context.Display(ais_sphere, False)
                        
                        # 存储AIS对象，用于后续访问
                        self.ais_shapes[node_shape_id] = ais_sphere
                    except Exception as e:
                        logger.error(f"绘制节点 {i} 时出错: {str(e)}")
                        logger.error(traceback.format_exc())
                
                logger.info(f"成功绘制 {len(self.node_shapes)} 个节点")
            
            # 更新显示
            self.viewer._display.View.Update()
            self.viewer._display.Repaint()
            self.viewer._display.FitAll()
            
            # 更新状态栏
            self.status_bar.showMessage(f"已绘制 {len(self.segment_shapes)} 条线段和 {len(self.node_shapes)} 个节点")
            logger.info(f"成功绘制 {len(self.segment_shapes)} 条线段和 {len(self.node_shapes)} 个节点")
        
        except Exception as e:
            logger.error(f"绘制线段时发生错误: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "绘制错误", f"绘制线段时出错: {str(e)}")

    def draw_imported_shapes(self, show_progress=True):
        """显示导入的STEP/IGES形状"""
        self.viewer._display.EraseAll()  # 清除现有显示
        self.ais_shapes = {}  # 清空AIS形状字典
        self.highlighted_shapes = []  # 清空高亮列表
        
        if not self.step_shapes:
            return
            
        # 显示进度对话框
        progress = None
        if show_progress:
            total_shapes = len(self.step_shapes)
            progress = QProgressDialog("正在显示导入的形状...", "取消", 0, total_shapes, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
            progress.show()
        
        # 批量处理形状
        shapes_batch = []
        batch_size = 50
        count = 0
        
        for shape_id, shape in self.step_shapes.items():
            # 更新进度条
            if progress:
                progress.setValue(count)
                QCoreApplication.processEvents()
                
                # 如果用户取消了操作
                if progress.wasCanceled():
                    break
                    
            count += 1
            
            try:
                # 使用AIS_Shape创建可视化对象
                ais_shape = AIS_Shape(shape)
                shapes_batch.append((shape_id, ais_shape))
                
                # 如果达到批处理大小或是最后一个元素，处理当前批次
                if len(shapes_batch) >= batch_size or count == len(self.step_shapes):
                    # 批量设置颜色和显示
                    for id, shape in shapes_batch:
                        # 设置默认颜色为蓝色
                        color = Quantity_Color(Quantity_NOC_BLUE)
                        self.viewer._display.Context.SetColor(shape, color, False)
                        
                        # 显示形状
                        self.viewer._display.Context.Display(shape, False)
                        
                        # 存储AIS对象
                        self.ais_shapes[id] = shape
                    
                    # 清空批处理列表
                    shapes_batch = []
                    # 更新视图
                    self.viewer._display.Context.UpdateCurrentViewer()
            
            except Exception as e:
                print(f"无法显示形状 {shape_id}: {e}")
                continue
        
        # 关闭进度对话框并适配视图
        if progress:
            progress.setValue(len(self.step_shapes))
            
        self.viewer._display.FitAll()
        self.viewer._display.Repaint()

    def highlight_shapes(self, shape_ids, highlight=True):
        """高亮显示或取消高亮指定的形状"""
        try:
            # 确保shape_ids是列表
            if not isinstance(shape_ids, list):
                shape_ids = [shape_ids]
                
            logger.debug(f"{'高亮' if highlight else '取消高亮'} {len(shape_ids)} 个形状: {shape_ids[:5]}...")
            
            needs_update = False
            for shape_id in shape_ids:
                if shape_id in self.ais_shapes:
                    ais_obj = self.ais_shapes[shape_id]
                    try:
                        if highlight:
                            # 根据形状类型设置不同的高亮颜色
                            if isinstance(shape_id, str) and shape_id.startswith('node_'):
                                color = Quantity_Color(Quantity_NOC_GREEN)
                            else: # 线段或导入形状
                                color = Quantity_Color(Quantity_NOC_YELLOW)
                            self.context.SetColor(ais_obj, color, False)
                        else:
                            # 恢复默认颜色
                            if isinstance(shape_id, str) and shape_id.startswith('node_'):
                                color = Quantity_Color(Quantity_NOC_RED) # 默认节点颜色
                            elif isinstance(shape_id, int): # 线段
                                color = Quantity_Color(Quantity_NOC_BLUE) # 默认线段颜色
                            else: # 导入形状
                                color = Quantity_Color(0.8, 0.8, 0.8, Quantity_TOC_RGB) # 默认导入颜色
                            self.context.SetColor(ais_obj, color, False)
                            
                        needs_update = True # 标记需要更新视图
                    except Exception as e:
                        logger.error(f"设置形状 {shape_id} 的高亮/颜色时出错: {e}")
                        continue
                else:
                    logger.warning(f"尝试高亮/取消高亮时找不到 shape_id: {shape_id}")
                    
            # 如果有更改，只更新一次视图
            if needs_update:
                try:
                    self.context.UpdateCurrentViewer()
                except Exception as e:
                    logger.error(f"更新视图以应用高亮时出错: {e}")
                    
        except Exception as e:
            logger.error(f"高亮形状时出错: {str(e)}")
            logger.error(traceback.format_exc())

    def on_tree_item_clicked(self, item, column):
        """处理树中项目点击事件"""
        self.selected_item = item
        if not item:
            return

        # 清除之前的高亮
        if hasattr(self, 'highlighted_shapes') and self.highlighted_shapes:
            self.highlight_shapes(self.highlighted_shapes, False)
        
        # 清空信息面板
        self.clear_info()
        
        # 获取项目数据
        data = item.data(0, Qt.UserRole)
        if data is None:
            return
        
        item_type = data.get("type")
        
        # 处理TotalNetwork点击 - 显示/隐藏所有子网络
        if item_type == "total_network":
            # 检查是否已有TotalNetwork相关的形状ID集合
            if hasattr(self, 'total_network_shapes') and self.total_network_shapes:
                # 检查第一个形状的可见性来确定当前状态
                first_shape = self.total_network_shapes[0] if self.total_network_shapes else None
                if first_shape:
                    # 获取当前可见性状态
                    current_visibility = not self.viewer._display.IsDisplayed(first_shape)
                    
                    # 显示或隐藏所有相关形状
                    for shape_id in self.total_network_shapes:
                        # 获取与ID对应的AIS_Shape对象
                        if shape_id in self.ais_shapes:
                            ais_obj = self.ais_shapes[shape_id]
                            if current_visibility:
                                self.context.Display(ais_obj, False)  # 显示，不立即更新
                            else:
                                self.context.Erase(ais_obj, False)    # 隐藏，不立即更新
                    
                    self.viewer._display.Repaint()
                    
                    # 更新状态栏
                    status = "显示" if current_visibility else "隐藏"
                    self.status_bar.showMessage(f"已{status} TotalNetwork 下的所有线路")
                    
                    # 更新信息面板
                    self.info_text.setPlainText(f"TotalNetwork\n已{status}所有子网络\n包含 {len(self.total_network_shapes)} 条线路")
                    
                    return
        
        # 处理network点击
        elif item_type == "network":
            network_idx = data.get("index")
            if network_idx is not None and network_idx in self.link_data:
                self.display_link_info(network_idx)
                
                # 如果在segment_shapes中找到对应的形状，则高亮显示
                if network_idx < len(self.segment_shapes):
                    shape_id = self.segment_shapes[network_idx]
                    self.highlight_shapes([shape_id])
        
        # 处理device或isopt点击，显示节点信息
        elif item_type in ["device", "isopt"]:
            node_name = data.get("name")
            if node_name in self.node_id_map:
                node_idx = self.node_id_map[node_name]
                shape_id = self.node_shapes[node_idx]
                self.display_node_info(shape_id)
                self.highlight_shapes([shape_id])

    def display_node_info(self, node_shape_id):
        """显示节点信息"""
        try:
            if node_shape_id not in self.shape_to_info:
                logger.warning(f"节点信息中找不到节点ID: {node_shape_id}")
                self.info_text.setText(f"找不到节点信息: {node_shape_id}")
                return

            node_info = self.shape_to_info[node_shape_id]
            node_name = node_info.get('name', '未知Name')
            coords = node_info.get('coordinates', (0,0,0))
            connected_links = node_info.get('connected_links', [])

            # 构建信息文本
            info_text = f"节点名称: {node_name}\n"
            info_text += f"(内部 ID: {node_shape_id})\n"
            info_text += f"坐标: X={coords[0]:.2f}, Y={coords[1]:.2f}, Z={coords[2]:.2f}\n"
            info_text += f"连接的链接数: {len(connected_links)}\n\n"

            # 添加连接的链接信息
            if connected_links:
                info_text += "连接的链接:\n"
                for link_idx in connected_links[:10]:  # 最多显示前10个
                    if link_idx in self.link_data:
                        link = self.link_data[link_idx]
                        info_text += f"- {link.get('name', f'链接_{link_idx}')} (ID: {link_idx})\n"

                if len(connected_links) > 10:
                    info_text += f"... 以及其他 {len(connected_links) - 10} 个链接\n"

            # 更新信息显示
            self.info_text.setText(info_text)

            # 更新状态栏
            self.status_bar.showMessage(f"已选择节点: {node_name}")

            logger.info(f"显示节点信息: {node_name} (ID: {node_shape_id})")
            logger.debug(f"节点坐标: X={coords[0]}, Y={coords[1]}, Z={coords[2]}")
            logger.debug(f"连接的链接数: {len(connected_links)}")
        except Exception as e:
            logger.error(f"显示节点信息时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"显示节点信息时出错: {str(e)}")
            
    def display_link_info(self, link_idx):
        """显示链接信息"""
        try:
            if link_idx not in self.link_data:
                logger.warning(f"链接数据中找不到索引: {link_idx}")
                self.info_text.setText(f"找不到链接信息: {link_idx}")
                return

            link = self.link_data[link_idx]

            # 构建信息文本
            info_text = f"链接: {link.get('name', f'链接_{link_idx}')}\n"
            info_text += f"(内部 ID: {link_idx})\n\n"

            # 起始点信息
            if 'start' in link:
                start = link['start']
                info_text += f"起始点名称: {start.get('name', '未知')}\n"
                coords = start.get('coordinates', (0,0,0))
                info_text += f"坐标: X={coords[0]:.2f}, Y={coords[1]:.2f}, Z={coords[2]:.2f}\n\n"

            # 终止点信息
            if 'end' in link:
                end = link['end']
                info_text += f"终止点名称: {end.get('name', '未知')}\n"
                coords = end.get('coordinates', (0,0,0))
                info_text += f"坐标: X={coords[0]:.2f}, Y={coords[1]:.2f}, Z={coords[2]:.2f}\n\n"

            # 其他信息
            for key, value in link.items():
                if key not in ['name', 'start', 'end', 'coordinates'] and not isinstance(value, dict):
                    info_text += f"{key}: {value}\n"

            # 更新信息显示
            self.info_text.setText(info_text)

            # 更新状态栏
            self.status_bar.showMessage(f"已选择链接: {link.get('name', f'链接_{link_idx}')}")

            logger.info(f"显示链接信息: {link.get('name', f'链接_{link_idx}')} (ID: {link_idx})")
        except Exception as e:
            logger.error(f"显示链接信息时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"显示链接信息时出错: {str(e)}")
            
    def clear_info(self):
        """清除显示的信息"""
        self.info_text.clear()
        self.status_bar.showMessage("点击3D形状以显示详细信息")
        
    def find_and_select_tree_item(self, shape_id):
        """根据形状ID查找并选择树中对应的项"""
        try:
            # 使用迭代方法避免深度递归问题
            items_to_check = []
            root = self.tree.invisibleRootItem()
            for i in range(root.childCount()):
                items_to_check.append(root.child(i))

            found_item = None
            while items_to_check:
                current_item = items_to_check.pop(0)  # 像队列一样处理
                item_data = current_item.data(0, Qt.UserRole)

                match = False
                if item_data == shape_id:
                    match = True
                elif isinstance(item_data, list) and shape_id in item_data:
                     # 如果是组节点，选择包含该ID的组
                     match = True

                if match:
                    found_item = current_item
                    break # 找到了

                # 添加子项到检查列表
                for i in range(current_item.childCount()):
                    items_to_check.append(current_item.child(i))

            if found_item:
                # 选择并滚动到项
                self.tree.setCurrentItem(found_item)
                self.tree.scrollToItem(found_item, QTreeWidget.PositionAtCenter)
                # 确保所有父项展开
                parent = found_item.parent()
                while parent and parent != self.tree.invisibleRootItem():
                    self.tree.expandItem(parent)
                    parent = parent.parent()
                logger.debug(f"在树中找到并选择了项: {found_item.text(0)} for ID {shape_id}")
            else:
                logger.debug(f"未找到与形状ID {shape_id} 对应的树项")
        except Exception as e:
            logger.error(f"查找和选择树项时出错: {str(e)}")
            logger.error(traceback.format_exc())
            
    def get_shape_type_name(self, shape):
        """获取形状类型的用户友好名称"""
        from OCC.Core.TopAbs import (
            TopAbs_COMPOUND, TopAbs_COMPSOLID, TopAbs_SOLID,
            TopAbs_SHELL, TopAbs_FACE, TopAbs_WIRE, TopAbs_EDGE, TopAbs_VERTEX
        )
        try:
            st = shape.ShapeType()
            if st == TopAbs_COMPOUND: return "复合体 (Compound)"
            if st == TopAbs_COMPSOLID: return "复合实体 (CompSolid)"
            if st == TopAbs_SOLID: return "实体 (Solid)"
            if st == TopAbs_SHELL: return "壳 (Shell)"
            if st == TopAbs_FACE: return "面 (Face)"
            if st == TopAbs_WIRE: return "线框 (Wire)"
            if st == TopAbs_EDGE: return "边 (Edge)"
            if st == TopAbs_VERTEX: return "顶点 (Vertex)"
            return f"形状 (Type {st})"
        except Exception as e:
            return "未知形状"

    def setup_interaction(self):
        """设置3D视图的交互功能"""
        try:
            # 注册鼠标点击事件回调
            self.viewer._display.register_select_callback(self.shape_selection_callback)

            # 设置状态栏初始信息
            self.status_bar.showMessage("点击3D形状以显示详细信息")
            logger.info("3D视图交互功能设置完成")
        except Exception as e:
            logger.error(f"设置交互功能时出错: {str(e)}")
            logger.error(traceback.format_exc())
            
    def shape_selection_callback(self, shape_list, *args):
        """当用户在3D视图中点击形状时的回调函数"""
        try:
            logger.debug(f"选择回调被触发，shape_list类型: {type(shape_list)}, 内容: {shape_list}")
            
            if not shape_list:
                logger.debug("没有选中任何形状")
                # 清除选择/信息
                if self.highlighted_shapes:
                   self.highlight_shapes(self.highlighted_shapes, False)
                   self.highlighted_shapes = []
                   self.clear_info()
                return

            # 通常我们关心第一个选择的形状
            selected_shape = shape_list[0]
            if not isinstance(selected_shape, TopoDS_Shape) or selected_shape.IsNull():
                logger.warning(f"选择回调收到的不是有效的TopoDS_Shape: {type(selected_shape)}")
                return

            logger.debug(f"选中的TopoDS_Shape类型: {self.get_shape_type_name(selected_shape)}")

            # 通过比较TopoDS_Shapes找到shape_id
            found_shape_id = None
            for shape_id, ais_obj in self.ais_shapes.items():
                try:
                    # 获取AIS_Shape中的TopoDS_Shape
                    ais_topo_shape = ais_obj.Shape()
                    if not ais_topo_shape.IsNull():
                        # 使用IsSame()进行比较
                        if ais_topo_shape.IsSame(selected_shape):
                            found_shape_id = shape_id
                            logger.debug(f"找到匹配的形状ID: {found_shape_id}")
                            break
                except Exception as e:
                    logger.warning(f"比较形状 {shape_id} 时出错: {e}")
                    continue

            if found_shape_id is None:
                logger.debug("未找到与选择形状匹配的ID")
                self.status_bar.showMessage("已选择形状，但无法在内部映射中找到它")
                if self.highlighted_shapes:
                   self.highlight_shapes(self.highlighted_shapes, False)
                   self.highlighted_shapes = []
                   self.clear_info()
                return

            # 在应用新高亮前清除之前的高亮
            if self.highlighted_shapes:
                # 避免在再次点击相同形状时取消高亮
                if self.highlighted_shapes != [found_shape_id]:
                    self.highlight_shapes(self.highlighted_shapes, False)

            # 高亮当前选择
            self.highlight_shapes([found_shape_id], True)
            self.highlighted_shapes = [found_shape_id] # 存储找到的ID

            # 根据ID类型显示信息
            if isinstance(found_shape_id, str) and found_shape_id.startswith('node_'):
                self.display_node_info(found_shape_id)
            elif isinstance(found_shape_id, int): # 线段使用整数索引
                self.display_link_info(found_shape_id)
            elif isinstance(found_shape_id, str): # 导入的STEP/IGES形状使用字符串ID
                logger.info(f"选中了导入的形状: {found_shape_id}")
                shape_info = self.shape_to_info.get(found_shape_id, {'type': 'Imported Shape'})
                info_text = f"选中的形状 ID: {found_shape_id}\n"
                info_text += f"类型: {shape_info.get('type', '未知')}"
                # 添加更多导入形状的详细信息
                self.info_text.setText(info_text)
                self.status_bar.showMessage(f"已选择导入的形状: {found_shape_id}")
            else:
                 logger.warning(f"未知的shape_id类型: {type(found_shape_id)}")
                 self.clear_info()
                 self.status_bar.showMessage(f"选择了未知类型的形状: {found_shape_id}")

            # 查找并选择树中对应的项
            self.find_and_select_tree_item(found_shape_id)

        except Exception as e:
            logger.error(f"形状选择回调中出现错误: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"选择处理错误: {str(e)}")

    def closeEvent(self, event):
        """窗口关闭时清理资源"""
        try:
            logger.info("正在关闭窗口，清理资源...")

            # 1. 清除所有显示的图形
            if hasattr(self, 'context') and self.context:
                try:
                    self.context.EraseAll(True)  # 立即清除所有显示
                    logger.debug("已清除所有显示的图形")
                except Exception as e:
                    logger.error(f"清除图形时出错: {e}")

            # 2. 关闭并释放 viewer 对象
            if hasattr(self, 'viewer') and self.viewer:
                try:
                    self.viewer.close()  # 关闭viewer
                    logger.debug("已关闭viewer")
                    self.viewer.deleteLater()  # 安排删除
                    logger.debug("已安排删除viewer")
                    self.viewer = None
                    self.context = None # 清除context引用
                except Exception as e:
                    logger.error(f"关闭viewer时出错: {e}")

            # 3. 清理计时器
            if hasattr(self, 'timer') and self.timer:
                try:
                    self.timer.stop()
                    logger.debug("已停止计时器")
                    if hasattr(self, 'timer_connected') and self.timer_connected:
                        try:
                            # 安全断开连接
                            self.timer.timeout.disconnect(self.close_message_box)
                            logger.debug("已断开计时器连接")
                        except Exception as e:
                            logger.warning(f"断开计时器连接时出错: {e}")
                        self.timer_connected = False
                except Exception as e:
                    logger.error(f"清理计时器时出错: {e}")

            logger.info("资源清理完成")
        except Exception as e:
            logger.error(f"清理资源时发生错误: {e}")
            logger.error(traceback.format_exc())
        finally:
            logger.info("调用父类closeEvent")
            super().closeEvent(event)
            
    def import_file(self):
        """根据选择的格式导入文件"""
        file_format = self.file_format_combo.currentText()
        logger.info(f"用户选择导入文件格式: {file_format}")
        
        if file_format == "STEP":
            self.import_step()
        elif file_format == "IGES":
            self.import_iges()
        else:
            logger.warning(f"不支持的文件格式: {file_format}")
            QMessageBox.warning(self, "格式错误", f"不支持的文件格式: {file_format}")
            
    def show_success_message(self, message):
        """显示成功消息，并在1秒后自动关闭"""
        try:
            # 如果已有消息框正在显示，先关闭它
            if hasattr(self, 'msg_box') and self.msg_box and self.msg_box.isVisible():
                self.msg_box.close()
                # 确保计时器停止
                if hasattr(self, 'timer') and self.timer.isActive():
                    self.timer.stop()
                    if hasattr(self, 'timer_connected') and self.timer_connected:
                        try:
                            self.timer.timeout.disconnect(self.close_message_box)
                        except Exception:
                            pass
                        self.timer_connected = False
                
            # 创建并显示新的消息框
            self.msg_box = QMessageBox(self)
            self.msg_box.setWindowTitle("成功")
            self.msg_box.setText(message)
            self.msg_box.setIcon(QMessageBox.Information)
            self.msg_box.setStandardButtons(QMessageBox.NoButton) # 隐藏按钮
            # 确保它是非模态的，不阻塞主窗口
            self.msg_box.setWindowModality(Qt.NonModal)
            self.msg_box.show()
            
            # 设置定时器自动关闭消息框
            if hasattr(self, 'timer'):
                if hasattr(self, 'timer_connected') and self.timer_connected:
                    try:
                        self.timer.timeout.disconnect(self.close_message_box)
                    except Exception:
                        pass
                # 连接并启动计时器
                self.timer.timeout.connect(self.close_message_box)
                self.timer_connected = True
                self.timer.start(1500)  # 1500毫秒 = 1.5秒
                
            logger.info(f"显示成功消息: {message}")
        except Exception as e:
            logger.error(f"显示成功消息时出错: {e}")
            logger.error(traceback.format_exc())
        
    def close_message_box(self):
        """关闭当前显示的消息框"""
        try:
            if hasattr(self, 'msg_box') and self.msg_box and self.msg_box.isVisible():
                self.msg_box.close()
                logger.debug("自动关闭成功消息框")
            
            # 断开计时器信号
            if hasattr(self, 'timer_connected') and self.timer_connected:
                try:
                    self.timer.timeout.disconnect(self.close_message_box)
                except Exception:
                    pass
                self.timer_connected = False
        except Exception as e:
            logger.error(f"关闭消息框时出错: {e}")
            logger.error(traceback.format_exc())

    def export_file(self):
        """根据选择的格式导出文件"""
        file_format = self.export_format_combo.currentText()
        logger.info(f"用户选择导出文件格式: {file_format}")
        
        # 确定要导出的内容
        shapes_to_export = []
        export_description = ""
        
        if self.main_shape and not self.main_shape.IsNull():
            shapes_to_export.append(self.main_shape)
            export_description = "主导入模型"
            logger.info("准备导出主导入模型...")
        elif self.step_shapes:
            # 导出所有单独导入的形状作为一个复合体
            shapes_to_export.extend(self.step_shapes.values())
            export_description = f"{len(shapes_to_export)} 个导入的形状"
            logger.info(f"准备导出 {len(shapes_to_export)} 个导入的形状...")
        elif self.segment_shapes or self.node_shapes:
            # 导出生成的线段和节点
            shapes_to_export.extend(self.segment_shapes)
            shapes_to_export.extend(self.node_shapes)
            export_description = f"{len(self.segment_shapes)} 条线段和 {len(self.node_shapes)} 个节点"
            logger.info(f"准备导出 {len(self.segment_shapes)} 条线段和 {len(self.node_shapes)} 个节点...")
        else:
            logger.warning("没有可导出的形状")
            QMessageBox.information(self, "无内容", "当前没有可导出的3D模型。请先导入文件或加载XML数据。")
            return
            
        # 过滤掉空形状
        valid_shapes_to_export = [s for s in shapes_to_export if s and not s.IsNull()]
        if not valid_shapes_to_export:
            logger.warning("过滤后没有有效的形状可导出")
            QMessageBox.warning(self, "无有效内容", "没有找到有效的形状进行导出。")
            return
            
        logger.info(f"将导出 {len(valid_shapes_to_export)} 个有效形状 ({export_description})")
        
        if file_format == "STEP":
            self.export_to_step(valid_shapes_to_export, export_description)
        elif file_format == "IGES":
            self.export_to_iges(valid_shapes_to_export, export_description)
        else:
            logger.warning(f"不支持的文件格式: {file_format}")
            QMessageBox.warning(self, "格式错误", f"不支持的文件格式: {file_format}")
            
    def export_to_step(self, shapes, description):
        """导出提供的形状列表为 STEP 文件"""
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, f"保存 {description} 为 STEP 文件", "", "STEP 文件 (*.step *.stp)")
            if not file_path:
                logger.info("用户取消了STEP导出")
                return
                
            logger.info(f"导出STEP文件到: {file_path}")
            
            # 显示进度对话框
            progress = QProgressDialog(f"正在导出 {description} 到 STEP...", "取消", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(500)
            progress.show()
            progress.setValue(10)
            QCoreApplication.processEvents()
            
            try:
                step_writer = STEPControl_Writer()
                Interface_Static_SetCVal("write.step.schema", "AP203") # 或 AP214
                
                progress.setValue(30)
                progress.setLabelText("正在处理形状...")
                QCoreApplication.processEvents()
                
                # 传输形状 - 如果是单个主形状，直接传输
                # 如果是多个形状，创建一个复合体
                if len(shapes) == 1 and shapes[0] == self.main_shape:
                    logger.info("正在传输主形状...")
                    transfer_ok = step_writer.Transfer(shapes[0], STEPControl_AsIs)
                    if not transfer_ok:
                        logger.error("传输主形状到 STEP writer 失败")
                        raise RuntimeError("无法传输主形状")
                else:
                    logger.info(f"正在创建包含 {len(shapes)} 个形状的复合体...")
                    compound = TopoDS_Compound()
                    builder = BRep_Builder()
                    builder.MakeCompound(compound)
                    added_count = 0
                    for i, shape in enumerate(shapes):
                        if progress.wasCanceled(): raise InterruptedError("用户取消导出")
                        try:
                            builder.Add(compound, shape)
                            added_count += 1
                            prog_val = 30 + int(40 * (i + 1) / len(shapes))
                            progress.setValue(prog_val)
                            QCoreApplication.processEvents()
                        except Exception as e:
                            logger.warning(f"添加形状 {i} 到复合体时出错: {e}")
                            continue # 跳过有问题的形状
                            
                    logger.info(f"成功添加 {added_count} 个形状到复合体。正在传输...")
                    progress.setValue(75)
                    progress.setLabelText("正在传输复合体...")
                    QCoreApplication.processEvents()
                    if added_count > 0:
                        transfer_ok = step_writer.Transfer(compound, STEPControl_AsIs)
                        if not transfer_ok:
                            logger.error("传输复合体到 STEP writer 失败")
                            raise RuntimeError("无法传输复合体")
                    else:
                        logger.error("没有有效的形状添加到复合体中")
                        raise RuntimeError("没有形状可导出")
                        
                progress.setValue(80)
                progress.setLabelText("正在写入文件...")
                QCoreApplication.processEvents()
                if progress.wasCanceled(): raise InterruptedError("用户取消导出")
                
                # 写入文件
                status = step_writer.Write(file_path)
                
                progress.setValue(100)
                
                if status == IFSelect_RetDone:
                    logger.info(f"STEP文件导出成功: {file_path}")
                    self.show_success_message(f"已成功导出到 {file_path}")
                else:
                    logger.error(f"STEP文件导出失败，状态: {status}")
                    QMessageBox.warning(self, "导出失败", f"导出过程中出现错误 (状态码: {status})，请检查日志。")
                    
            except InterruptedError:
                logger.info("用户取消了STEP导出")
                progress.close()
            except Exception as e:
                progress.close()
                logger.error(f"导出STEP文件时发生异常: {e}")
                logger.error(traceback.format_exc())
                QMessageBox.critical(self, "导出错误", f"导出过程中出现异常: {str(e)}")
        except Exception as e:
            logger.error(f"STEP导出过程中发生未捕获异常: {e}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "导出错误", f"导出过程中发生未预期的错误: {str(e)}")
            
    def export_to_iges(self, shapes, description):
        """导出提供的形状列表为 IGES 文件"""
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, f"保存 {description} 为 IGES 文件", "", "IGES 文件 (*.igs *.iges)")
            if not file_path:
                logger.info("用户取消了IGES导出")
                return
                
            logger.info(f"导出IGES文件到: {file_path}")
            
            progress = QProgressDialog(f"正在导出 {description} 到 IGES...", "取消", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(500)
            progress.show()
            progress.setValue(10)
            QCoreApplication.processEvents()
            
            try:
                iges_writer = IGESControl_Writer()
                # 设置IGES参数（如果需要）
                # Interface_Static_SetCVal("write.iges.unit", "MM")
                # Interface_Static_SetCVal("write.iges.brep.mode", "1") # B-Rep solid
                
                progress.setValue(30)
                progress.setLabelText("正在添加形状...")
                QCoreApplication.processEvents()
                
                added_count = 0
                for i, shape in enumerate(shapes):
                    if progress.wasCanceled(): raise InterruptedError("用户取消导出")
                    try:
                        add_ok = iges_writer.AddShape(shape)
                        if not add_ok:
                            logger.warning(f"无法将形状 {i} 添加到 IGES writer")
                        else:
                            added_count += 1
                        prog_val = 30 + int(50 * (i + 1) / len(shapes))
                        progress.setValue(prog_val)
                        QCoreApplication.processEvents()
                    except Exception as e:
                        logger.warning(f"添加形状 {i} 到 IGES writer 时出错: {e}")
                        continue
                        
                if added_count == 0:
                    logger.error("没有形状成功添加到 IGES writer")
                    raise RuntimeError("没有形状可导出到 IGES")
                    
                logger.info(f"成功添加 {added_count} 个形状。正在写入文件...")
                progress.setValue(85)
                progress.setLabelText("正在写入文件...")
                QCoreApplication.processEvents()
                if progress.wasCanceled(): raise InterruptedError("用户取消导出")
                
                # 写入文件
                write_ok = iges_writer.Write(file_path)
                
                progress.setValue(100)
                
                if write_ok:
                    logger.info(f"IGES文件导出成功: {file_path}")
                    self.show_success_message(f"已成功导出到 {file_path}")
                else:
                    # IGES Write 返回布尔值，而不是像STEP那样的状态码
                    logger.error("IGES文件导出失败 (Write 方法返回 False)")
                    QMessageBox.warning(self, "导出失败", "导出 IGES 文件过程中出现错误，请检查日志。")
                    
            except InterruptedError:
                logger.info("用户取消了IGES导出")
                progress.close()
            except Exception as e:
                progress.close()
                logger.error(f"导出IGES文件时发生异常: {e}")
                logger.error(traceback.format_exc())
                QMessageBox.critical(self, "导出错误", f"导出过程中出现异常: {str(e)}")
        except Exception as e:
            logger.error(f"IGES导出过程中发生未捕获异常: {e}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "导出错误", f"导出过程中发生未预期的错误: {str(e)}")
            
    def import_step(self):
        """导入STEP文件并解析"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 STEP 文件", "", "STEP 文件 (*.step *.stp)")
        if not file_path:
            logger.info("用户取消了STEP导入")
            return
            
        logger.info(f"选择导入STEP文件: {file_path}")
        self.import_cad_file(file_path, "STEP")
        
    def import_iges(self):
        """导入IGES文件并解析"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 IGES 文件", "", "IGES 文件 (*.igs *.iges)")
        if not file_path:
            logger.info("用户取消了IGES导入")
            return
            
        logger.info(f"选择导入IGES文件: {file_path}")
        self.import_cad_file(file_path, "IGES")
        
    def import_cad_file(self, file_path, file_format):
        """通用CAD文件导入函数"""
        logger.info(f"开始导入 {file_format} 文件: {file_path}")
        progress = QProgressDialog(f"正在导入 {file_format} 文件...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(500)
        progress.show()
        progress.setValue(5)
        QCoreApplication.processEvents()
        
        try:
            # 清除现有数据
            logger.info("清除现有数据...")
            self.tree.clear()
            self.segment_shapes = []
            self.segments = []
            self.node_shapes = []
            self.unique_nodes = {}
            self.node_id_map = {}
            self.link_data = {}
            self.node_to_links = {}
            self.shape_to_info = {}
            self.step_shapes = {} # 清除先前导入的形状
            self.ais_shapes = {}  # 清除AIS对象
            self.highlighted_shapes = []
            self.main_shape = None
            self.context.EraseAll(True) # 清除视图
            logger.info("现有数据清除完毕")
            progress.setValue(10)
            QCoreApplication.processEvents()
            
            # 读取文件
            logger.info(f"使用 {file_format} reader 读取文件...")
            if file_format == "STEP":
                reader = STEPControl_Reader()
            elif file_format == "IGES":
                reader = IGESControl_Reader()
            else:
                raise ValueError(f"不支持的文件格式: {file_format}")
                
            read_status = reader.ReadFile(file_path)
            logger.info(f"ReadFile 状态: {read_status}")
            progress.setValue(40)
            QCoreApplication.processEvents()
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")
            
            if read_status != IFSelect_RetDone:
                # 尝试获取更多错误信息（可能不总是有效）
                fail_messages = reader.FailMessage() if hasattr(reader, 'FailMessage') else ""
                error_msg = f"读取{file_format}文件失败 (状态: {read_status})"
                if fail_messages:
                    error_msg += f"\n错误信息:\n{fail_messages}"
                    logger.error(f"{error_msg}")
                else:
                    logger.error(error_msg)
                raise RuntimeError(error_msg)
                
            # 转换根
            logger.info("转换文件根...")
            # 检查根数量
            num_roots = reader.NbRootsForTransfer()
            logger.info(f"文件中根的数量: {num_roots}")
            if num_roots == 0:
                raise RuntimeError(f"{file_format} 文件中没有找到可转换的根。文件可能为空或格式无效。")
                
            transfer_ok = reader.TransferRoots()
            logger.info(f"TransferRoots 状态: {transfer_ok}")
            progress.setValue(60)
            QCoreApplication.processEvents()
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")
            if not transfer_ok:
                raise RuntimeError(f"转换{file_format}文件根失败。")
                
            # 获取形状
            logger.info("获取主形状...")
            # 检查结果形状数量
            num_shapes = reader.NbShapes()
            logger.info(f"转换后的形状数量: {num_shapes}")
            if num_shapes == 0:
                raise RuntimeError(f"转换{file_format}文件后未生成任何形状。")
            elif num_shapes == 1:
                logger.info("获取单个主形状...")
                shape = reader.Shape(1) # 获取第一个形状
            else:
                # 如果有多个形状，创建一个复合体
                logger.info(f"获取 {num_shapes} 个形状并创建复合体...")
                shape = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(shape)
                for i in range(1, num_shapes + 1):
                    sub_shape = reader.Shape(i)
                    if sub_shape and not sub_shape.IsNull():
                        builder.Add(shape, sub_shape)
                    else:
                        logger.warning(f"跳过转换结果中的空形状索引 {i}")
                        
            if shape.IsNull():
                raise RuntimeError(f"从{file_format}文件获取的最终形状为空。")
                
            self.main_shape = shape # 存储主形状
            logger.info(f"主形状类型: {self.get_shape_type_name(shape)}")
            progress.setValue(70)
            QCoreApplication.processEvents()
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")
            
            # 分析并构建树
            logger.info("分析形状并构建树...")
            file_basename = os.path.basename(file_path)
            root = QTreeWidgetItem(self.tree)
            root.setText(0, f"{file_format} Model: {file_basename}")
            
            # 传递进度对话框以允许在分析期间取消
            self.analyze_shape_and_build_tree(self.main_shape, root, progress)
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")
            
            self.tree.expandItem(root)
            progress.setValue(90)
            QCoreApplication.processEvents()
            
            # 绘制导入的形状
            logger.info("绘制导入的形状...")
            self.draw_imported_shapes(show_progress=True) # 绘制有自己的进度
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")
            
            progress.setValue(100)
            
            total_elements = len(self.step_shapes) # 计算单个形状数量
            logger.info(f"{file_format}文件导入成功，包含 {total_elements} 个子元素")
            self.show_success_message(f"已成功导入 {file_format} 文件 ({file_basename})，包含 {total_elements} 个元素。")
            
        except InterruptedError:
            logger.warning(f"用户取消了 {file_format} 导入")
            progress.close()
            # 清理可能部分加载的状态
            self.tree.clear()
            self.context.EraseAll(True)
            self.status_bar.showMessage(f"{file_format} 导入已取消")
        except Exception as e:
            progress.close()
            logger.error(f"导入{file_format}文件时发生异常: {e}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "导入错误", f"导入 {file_path} 时出现异常:\n\n{str(e)}\n\n请查看日志获取详细信息。")
            # 清理
            self.tree.clear()
            self.context.EraseAll(True)
            
    def analyze_shape_and_build_tree(self, shape, parent_item, progress=None):
        """分析形状的层次结构并构建树视图"""
        from OCC.Core.TopAbs import (
            TopAbs_COMPOUND, TopAbs_COMPSOLID, TopAbs_SOLID,
            TopAbs_SHELL, TopAbs_FACE, TopAbs_WIRE, TopAbs_EDGE, TopAbs_VERTEX
        )
        from OCC.Core.TopExp import TopExp_Explorer
        from OCC.Core.TopTools import TopTools_IndexedMapOfShape
        
        try:
            logger.info("开始分析形状...")
            if progress: progress.setLabelText("分析形状结构...")
            
            # 清除先前导入的形状
            self.step_shapes = {}
            all_shape_ids_in_tree = [] # 跟踪添加到树中的ID
            
            # 使用TopExp_Explorer遍历子形状
            # 我们主要关注实体、壳、面、边进行可视化/选择
            shape_types_to_explore = {
                TopAbs_SOLID: "实体 (Solid)",
                TopAbs_SHELL: "壳 (Shell)",
                TopAbs_FACE: "面 (Face)",
                # TopAbs_WIRE: "线框 (Wire)", # 通常是面的一部分，可能添加噪声
                TopAbs_EDGE: "边 (Edge)"
            }
            type_nodes = {} # 存储树中的顶级类型节点
            
            # 估计总工作量以更新进度
            # 预先计算形状可能对复杂模型很慢
            # 我们将基于迭代更新进度
            
            processed_count = 0
            unique_subshapes = TopTools_IndexedMapOfShape() # 使用映射避免重复
            
            # 第一遍：收集相关类型的唯一子形状
            logger.debug("第一遍: 收集唯一子形状...")
            for shape_type in shape_types_to_explore:
                explorer = TopExp_Explorer(shape, shape_type)
                while explorer.More():
                    current_sub = explorer.Current()
                    if not current_sub.IsNull():
                        unique_subshapes.Add(current_sub)
                    explorer.Next()
            total_unique = unique_subshapes.Extent()
            logger.debug(f"找到 {total_unique} 个唯一子形状 (相关类型)")
            if progress: progress.setLabelText(f"分析 {total_unique} 个形状...")
            
            # 第二遍：迭代唯一形状，创建树项，存储形状
            logger.debug("第二遍: 创建树节点并存储形状...")
            shape_counter = 0
            for i in range(1, total_unique + 1):
                if progress and progress.wasCanceled(): raise InterruptedError("用户取消分析")
                
                current_shape = unique_subshapes.FindKey(i)
                if current_shape.IsNull(): continue
                
                shape_type = current_shape.ShapeType()
                type_name = self.get_shape_type_name(current_shape) # 获取用户友好名称
                base_type_name = shape_types_to_explore.get(shape_type, "其他") # 获取类别名称
                
                # 获取或创建树中的类别节点
                if base_type_name not in type_nodes:
                    type_node = QTreeWidgetItem(parent_item)
                    type_node.setText(0, base_type_name)
                    type_nodes[base_type_name] = {"node": type_node, "ids": []}
                else:
                    type_node = type_nodes[base_type_name]["node"]
                    
                # 为单个形状创建项
                shape_item = QTreeWidgetItem(type_node)
                # 生成唯一ID: type_index_hash
                shape_id = f"{base_type_name.split(' ')[0].lower()}_{shape_counter}_{current_shape.HashCode(1000000)}"
                shape_item.setText(0, f"{type_name} ID: {shape_id}")
                shape_item.setData(0, Qt.UserRole, shape_id) # 存储ID
                
                # 在字典中存储实际TopoDS_Shape
                self.step_shapes[shape_id] = current_shape
                # 存储基本信息（稍后可以扩展）
                self.shape_to_info[shape_id] = {'type': type_name}
                
                # 将ID添加到组高亮列表
                type_nodes[base_type_name]["ids"].append(shape_id)
                all_shape_ids_in_tree.append(shape_id)
                shape_counter += 1
                
                processed_count += 1
                if progress and total_unique > 0:
                    prog_val = 70 + int(20 * processed_count / total_unique) # 进度从70%到90%
                    progress.setValue(prog_val)
                    QCoreApplication.processEvents()
                    
            # 更新类型节点上的计数并存储ID列表
            logger.debug("更新树节点计数和数据...")
            for base_type_name, data in type_nodes.items():
                node = data["node"]
                ids = data["ids"]
                node.setText(0, f"{base_type_name} ({len(ids)})")
                node.setData(0, Qt.UserRole, ids)
                
            # 同样在根项中存储所有收集的ID
            parent_item.setData(0, Qt.UserRole, all_shape_ids_in_tree)
            
            # 自身存储主形状（如果没有拆分，例如简单实体）
            # 我们已经在导入过程中存储了self.main_shape
            
            logger.info(f"形状分析完成，存储了 {len(self.step_shapes)} 个形状，创建了 {len(all_shape_ids_in_tree)} 个树条目")
            
        except InterruptedError:
            logger.warning("形状分析被用户取消")
            raise # 重新抛出以被导入器捕获
        except Exception as e:
            logger.error(f"分析形状和构建树时出错: {e}")
            logger.error(traceback.format_exc())
            # 不在此处抛出，让导入有机会完成但显示警告
            QMessageBox.warning(self, "分析警告", f"分析形状结构时出现错误:\n{e}\n导入结果可能不完整。")


def main(xml_file=None):
    # 设置日志系统
    log_file = setup_logging()
    
    # 创建 QApplication 实例
    app = QApplication.instance()  # 检查是否已存在
    if not app:  # 如果不存在，创建一个
        app = QApplication(sys.argv)
        
    logger.info(f"应用程序启动，日志文件: {log_file}")
    logger.info(f"Python版本: {sys.version}")
    logger.info(f"Qt版本: {QT_VERSION_STR}")
    logger.info(f"OCC Backend: qt-pyqt5")
    
    # 显示命令行参数
    logger.info(f"命令行参数: {sys.argv}")
    if xml_file:
        logger.info(f"指定的XML文件: {xml_file}")
    else:
        logger.info("未指定XML文件，将启动空窗口。")
        
    # 检查文件是否存在
    if xml_file and not os.path.exists(xml_file):
        logger.error(f"指定的XML文件不存在: {xml_file}")
        print(f"错误: 文件未找到 '{xml_file}'")
        # 显示错误消息
        error_msg = QMessageBox()
        error_msg.setIcon(QMessageBox.Critical)
        error_msg.setWindowTitle("文件错误")
        error_msg.setText(f"无法找到指定的XML文件:\n{xml_file}\n\n应用程序将以空状态启动。")
        error_msg.exec_()
        xml_file = None  # 进入空状态启动
        
    # 创建主窗口
    try:
        window = MainWindow(xml_file)
        
        # 设置窗口大小和居中显示
        window.resize(1200, 900)  # 稍大的默认尺寸
        try:
            screen_geometry = QDesktopWidget().availableGeometry()  # 使用可用屏幕空间
            x = (screen_geometry.width() - window.width()) // 2
            y = (screen_geometry.height() - window.height()) // 2
            window.move(x, y)
            logger.info(f"窗口大小设置为 {window.width()}x{window.height()}，居中显示在位置 ({x}, {y})")
        except Exception as e:
            logger.warning(f"无法获取屏幕几何信息或居中窗口: {e}")
            
        logger.info("显示主窗口")
        window.show()
        
        # 启动应用程序事件循环
        exit_code = app.exec_()
        logger.info(f"应用程序退出，退出代码: {exit_code}")
        return exit_code
    except Exception as e:
        logger.critical(f"创建或显示主窗口时发生未捕获异常: {e}")
        logger.critical(traceback.format_exc())
        # 显示最终严重错误消息
        error_msg = QMessageBox()
        error_msg.setIcon(QMessageBox.Critical)
        error_msg.setWindowTitle("应用程序错误")
        error_msg.setText(f"应用程序遇到严重错误并需要关闭:\n\n{str(e)}\n\n请查看日志文件 '{log_file}' 获取详细信息。")
        error_msg.exec_()
        return 1


if __name__ == "__main__":
    # 使用 argparse 解析命令行参数
    parser = argparse.ArgumentParser(description="航电布线可视化系统")
    parser.add_argument("xml_file", type=str, nargs='?', default=None, help="XML 文件路径")
    parser.add_argument("--debug", action="store_true", help="启用详细调试日志")
    args = parser.parse_args()
    
    # 根据debug标志配置日志级别
    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.getLogger().setLevel(log_level)
    
    if args.debug:
        print("--- 调试模式已启用 ---")
        
    try:
        # 调用主函数并使用其返回值退出
        exit_status = main(args.xml_file)
        sys.exit(exit_status)
    except Exception as e:
        # 捕获启动或关闭期间的任何意外异常
        print(f"发生未捕获的顶层异常: {e}")
        traceback.print_exc()
        # 尝试记录异常（如果日志系统可用）
        try:
            logger.critical(f"发生未捕获的顶层异常: {e}", exc_info=True)
        except:
            pass  # 日志系统可能未初始化
        sys.exit(1)