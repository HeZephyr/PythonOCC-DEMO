# -*- coding: utf-8 -*-
import sys
import os
import argparse
import logging
import traceback
from datetime import datetime
from OCC.Core.Quantity import Quantity_Color
from OCC.Core._Quantity import Quantity_TOC_RGB
from PyQt5.QtCore import QT_VERSION_STR
from OCC.Core.gp import gp_Pnt, gp_Vec, gp_Dir, gp_Ax2
from OCC.Display.backend import load_backend
from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_Reader, STEPControl_AsIs
from OCC.Core.IGESControl import IGESControl_Reader, IGESControl_Writer
from OCC.Core.Interface import Interface_Static_SetCVal
from OCC.Core.IFSelect import IFSelect_RetDone
from OCC.Core.BRep import BRep_Tool, BRep_Builder

from OCC.Core.V3d import V3d_Zneg, V3d_Yneg, V3d_Xneg

load_backend("qt-pyqt5")
from OCC.Display.qtDisplay import qtViewer3d
from PyQt5.QtWidgets import (
    QApplication,
    QTreeWidget,
    QTreeWidgetItem,
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QDesktopWidget,
    QPushButton,
    QMessageBox,
    QFileDialog,
    QProgressDialog,
    QComboBox,
    QLabel,
    QGroupBox,
    QTextEdit,
    QStatusBar
)
from OCC.Core.BRepPrimAPI import BRepPrimAPI_MakeCylinder, BRepPrimAPI_MakeSphere
import pandas as pd
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QTimer, QCoreApplication

# 添加以下导入用于STEP文件解析
from OCC.Core.TopoDS import TopoDS_Shape, TopoDS_Edge, topods_Edge, TopoDS_Compound, topods_Compound, topods
from OCC.Core.TopExp import TopExp_Explorer
from OCC.Core.TopAbs import (
    TopAbs_EDGE, TopAbs_VERTEX, TopAbs_FACE, TopAbs_SOLID,
    TopAbs_SHELL, TopAbs_COMPOUND, TopAbs_COMPSOLID, TopAbs_WIRE
)
from OCC.Core.TopLoc import TopLoc_Location
from OCC.Core.BRepAdaptor import BRepAdaptor_Curve
from OCC.Core.GeomAbs import GeomAbs_Line
from OCC.Core.TopTools import TopTools_IndexedMapOfShape
from OCC.Core.BRepTools import breptools_OuterWire
from OCC.Core.ShapeAnalysis import ShapeAnalysis_Edge
from OCC.Core.AIS import AIS_Shape, AIS_InteractiveContext # Import AIS_InteractiveContext
from OCC.Core.Quantity import Quantity_Color, Quantity_NOC_BLUE, Quantity_NOC_YELLOW, Quantity_NOC_RED, Quantity_NOC_GREEN

def setup_logging():
    """设置日志记录（修复版本）"""
    log_dir = "logs"
    try:
        os.makedirs(log_dir, exist_ok=True)
        if not os.access(log_dir, os.W_OK):
            raise PermissionError(f"无写入权限: {log_dir}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"app_{timestamp}.log")

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

        # 创建应用专属日志器（必须关闭传播防止重复记录）
        app_logger = logging.getLogger("visualize_xlsx")
        app_logger.propagate = False  # 关闭传播，避免被根日志器重复处理
        app_logger.setLevel(logging.DEBUG)
        # 直接添加处理器到应用日志器
        app_logger.addHandler(file_handler)
        app_logger.addHandler(console_handler)

        return log_file

    except Exception as e:
        print(f"致命错误: 无法初始化日志系统 - {str(e)}")
        traceback.print_exc()
        sys.exit(1)

# 全局日志器
logger = logging.getLogger("visualize_xlsx")

class MainWindow(QWidget):
    def __init__(self, df=None):
        super().__init__()
        self.setWindowTitle("基于公共数据源的航电系统布线架构与集成设计技术研究系统")

        logger.info("初始化应用程序...")

        # 初始化消息框和计时器
        self.msg_box = None
        self.timer = QTimer()
        self.timer.setSingleShot(True)

        # 标记计时器是否已连接
        self.timer_connected = False

        # 标记是否首次绘制
        self.first_draw = True

        # 高亮和颜色管理
        self.ais_shapes = {}  # 存储所有AIS对象，用于颜色管理 {shape_id: AIS_Shape}
        self.highlighted_shapes = []  # 当前高亮的形状IDs

        # 节点相关数据结构
        self.unique_nodes = {}  # 存储唯一节点信息，使用ref作为键，(x, y, z)作为值
        self.node_shapes = []  # 存储节点的 TopoDS_Shape 对象
        self.node_id_map = {} # Map node_ref to node_index ('node_0', 'node_1', etc.)

        # 存储链接相关的数据，用于查询
        self.link_data = {}  # 存储链接数据，使用索引作为键
        self.node_to_links = {}  # 存储节点关联的链接，使用节点ref作为键
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
        
         # 布局操作
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
        # self.context.SetBackground(bg_color)
        self.viewer._display.View.SetBackgroundColor(bg_color) # Deprecated way
        horizontal_layout.addWidget(self.viewer, 2)
        
         # 布局操作绑定
        self.layout_button.clicked.connect(self.viewer._display.FitAll)
        # 正视图
        self.front_view_button.clicked.connect(self.set_front_view)
        # 俯视图
        self.top_view_button.clicked.connect(self.set_top_view)
        # 右视图
        self.right_view_button.clicked.connect(self.set_right_view)

        # 添加主水平布局
        main_layout.addLayout(horizontal_layout)

        # 添加状态栏
        self.status_bar = QStatusBar()
        self.status_bar.setFixedHeight(25)
        main_layout.addWidget(self.status_bar)

        self.setLayout(main_layout)

        # 存储线段形状和原始颜色
        self.segment_shapes = [] # Store TopoDS_Shape for segments
        self.segments = [] # Store segment start/end points

        # 保存导入的模型
        self.step_shapes = {}  # 形状ID到形状对象的映射 {shape_id: TopoDS_Shape}
        self.main_shape = None # Store the main imported shape

        # 树项点击事件
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        self.selected_item = None

        # 设置交互功能
        self.setup_interaction()

        # 如果有数据，则解析并显示
        if df is not None:
            try:
                self.parse_df_and_populate_tree(df)
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

    def parse_df_and_populate_tree(self, df):
        """Parse dataframe and populate the tree widget with hierarchical structure."""
        logger.info(f"开始解析数据框，行数: {len(df)}")
        self.tree.clear()  # 清空树
        self.segment_shapes = []
        self.segments = []
        self.ais_shapes = {}  # 清空AIS形状字典
        self.step_shapes = {} # Clear imported shapes
        self.main_shape = None

        # 存储唯一节点信息
        self.unique_nodes = {}  # 使用ref作为键，(x, y, z)作为值
        self.node_shapes = []  # 存储节点的 TopoDS_Shape 对象
        self.node_id_map = {} # Clear node mapping

        # 初始化链接数据和节点到链接的映射
        self.link_data = {}
        self.node_to_links = {}
        self.shape_to_info = {} # Clear shape info mapping

        # 显示进度对话框
        progress = QProgressDialog("正在解析Excel数据...", "取消", 0, len(df), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
        progress.show()
        QCoreApplication.processEvents()

        main_root = QTreeWidgetItem(self.tree)
        main_root.setText(0, "NETWORK AIRPLANE")
        main_root.setData(0, Qt.UserRole, [])  # 将用于存储所有子节点的索引

        # 添加节点根节点
        nodes_root = QTreeWidgetItem(main_root)
        nodes_root.setText(0, "Network Nodes")
        nodes_root.setData(0, Qt.UserRole, [])  # 初始化节点索引列表

        section_groups = {}
        section_indices = {}  # 记录每个section包含的线段索引

        for index, row in df.iterrows():
            try:
                # 更新进度条
                progress.setValue(index)
                QCoreApplication.processEvents()

                # 如果用户取消了操作
                if progress.wasCanceled():
                    logger.info("用户取消了Excel数据解析")
                    break

                # 检查并安全获取各列数据
                try:
                    link_name = str(row.get('Link Name', f"Link_{index}"))
                    ref_origine = str(row.get('refOrigine', f"Origin_{index}"))
                    x_origine = float(row.get('Xorigine', 0.0))
                    y_origine = float(row.get('Yorigine', 0.0))
                    z_origine = float(row.get('Zorigine', 0.0))
                    ref_extremite = str(row.get('RefExtremite', f"Extremite_{index}"))
                    x_extremite = float(row.get('Xextremite', 0.0))
                    y_extremite = float(row.get('Yextremite', 0.0))
                    z_extremite = float(row.get('Zextremite', 0.0))
                    length = row.get('Length', 0.0)
                    density = row.get('Density', 0.0)
                    safety = row.get('Safety', '')
                    route = row.get('Route', '')
                    action_number = row.get('Action Number', '')
                    section = str(row.get('Section', 'Default'))
                except (ValueError, TypeError) as e:
                    logger.warning(f"在行 {index} 处理列数据时出错: {e}")
                    # 使用默认值继续
                    if 'link_name' not in locals(): link_name = f"Link_{index}"
                    if 'ref_origine' not in locals(): ref_origine = f"Origin_{index}"
                    if 'x_origine' not in locals(): x_origine = 0.0
                    if 'y_origine' not in locals(): y_origine = 0.0
                    if 'z_origine' not in locals(): z_origine = 0.0
                    if 'ref_extremite' not in locals(): ref_extremite = f"Extremite_{index}"
                    if 'x_extremite' not in locals(): x_extremite = 0.0
                    if 'y_extremite' not in locals(): y_extremite = 0.0
                    if 'z_extremite' not in locals(): z_extremite = 0.0
                    if 'length' not in locals(): length = 0.0
                    if 'density' not in locals(): density = 0.0
                    if 'safety' not in locals(): safety = ''
                    if 'route' not in locals(): route = ''
                    if 'action_number' not in locals(): action_number = ''
                    if 'section' not in locals(): section = 'Default'

                # 添加唯一节点信息
                if ref_origine not in self.unique_nodes:
                    self.unique_nodes[ref_origine] = (x_origine, y_origine, z_origine)

                if ref_extremite not in self.unique_nodes:
                    self.unique_nodes[ref_extremite] = (x_extremite, y_extremite, z_extremite)

                # 创建或获取section组
                if section not in section_groups:
                    section_group_item = QTreeWidgetItem(main_root)
                    section_group_item.setText(0, f"Network Geometry {section}")
                    section_groups[section] = section_group_item
                    section_indices[section] = []  # 初始化该section的索引列表
                else:
                    section_group_item = section_groups[section]

                link_item = QTreeWidgetItem(section_group_item)
                link_item.setText(0, link_name)

                # 添加 origin 信息
                origin_item = QTreeWidgetItem(link_item)
                origin_item.setText(0, "Origin")
                ref_origine_subitem = QTreeWidgetItem(origin_item)
                ref_origine_subitem.setText(0, f"refOrigine= {ref_origine}")
                x_origine_subitem = QTreeWidgetItem(origin_item)
                x_origine_subitem.setText(0, f"Xorigine= {x_origine}")
                y_origine_subitem = QTreeWidgetItem(origin_item)
                y_origine_subitem.setText(0, f"Yorigine= {y_origine}")
                z_origine_subitem = QTreeWidgetItem(origin_item)
                z_origine_subitem.setText(0, f"Zorigine= {z_origine}")

                # 添加 Extremite 信息
                extremite_item = QTreeWidgetItem(link_item)
                extremite_item.setText(0, "Extremite")
                ref_extremite_subitem = QTreeWidgetItem(extremite_item)
                ref_extremite_subitem.setText(0, f"RefExtremite= {ref_extremite}")
                x_extremite_subitem = QTreeWidgetItem(extremite_item)
                x_extremite_subitem.setText(0, f"Xextremite= {x_extremite}")
                y_extremite_subitem = QTreeWidgetItem(extremite_item)
                y_extremite_subitem.setText(0, f"Yextremite= {y_extremite}")
                z_extremite_subitem = QTreeWidgetItem(extremite_item)
                z_extremite_subitem.setText(0, f"Zextremite= {z_extremite}")

                # 添加其他信息
                other_info_item = QTreeWidgetItem(link_item)
                other_info_item.setText(0, "Other Info")
                length_subitem = QTreeWidgetItem(other_info_item)
                length_subitem.setText(0, f"Length= {length}")
                density_subitem = QTreeWidgetItem(other_info_item)
                density_subitem.setText(0, f"Density= {density}")
                safety_subitem = QTreeWidgetItem(other_info_item)
                safety_subitem.setText(0, f"Safety= {safety}")
                route_subitem = QTreeWidgetItem(other_info_item)
                route_subitem.setText(0, f"Route= {route}")
                action_number_subitem = QTreeWidgetItem(other_info_item)
                action_number_subitem.setText(0, f"Action Number= {action_number}")
                section_subitem = QTreeWidgetItem(other_info_item)
                section_subitem.setText(0, f"Section= {section}")

                start_point = (x_origine, y_origine, z_origine)
                end_point = (x_extremite, y_extremite, z_extremite)
                self.segments.append((start_point, end_point))

                # 存储当前线段的索引 (Use integer index as shape_id for segments)
                segment_index = len(self.segments) - 1
                link_item.setData(0, Qt.UserRole, segment_index)

                # 将索引添加到对应的section索引列表中
                section_indices[section].append(segment_index)

                # 存储链接数据，用于查询
                self.link_data[segment_index] = {
                    'name': link_name,
                    'origin': {
                        'ref': ref_origine,
                        'coordinates': (x_origine, y_origine, z_origine)
                    },
                    'extremite': {
                        'ref': ref_extremite,
                        'coordinates': (x_extremite, y_extremite, z_extremite)
                    },
                    'length': length,
                    'density': density,
                    'safety': safety,
                    'route': route,
                    'action_number': action_number,
                    'section': section
                }

                # 存储形状信息 (using segment index as key)
                self.shape_to_info[segment_index] = self.link_data[segment_index]

                # 更新节点到链接的映射 (using node ref as key)
                if ref_origine not in self.node_to_links:
                    self.node_to_links[ref_origine] = []
                self.node_to_links[ref_origine].append(segment_index)

                if ref_extremite not in self.node_to_links:
                    self.node_to_links[ref_extremite] = []
                self.node_to_links[ref_extremite].append(segment_index)
            except Exception as e:
                logger.error(f"处理行 {index} 时出错: {str(e)}")
                logger.error(traceback.format_exc())
                continue

        try:
            # 现在添加节点到树中
            node_indices_for_tree = [] # Store node shape_ids ('node_0', 'node_1'...)

            for i, (node_ref, node_pos) in enumerate(self.unique_nodes.items()):
                node_item = QTreeWidgetItem(nodes_root)
                node_item.setText(0, f"Node: {node_ref}")

                # 使用 'node_i' 格式作为 shape_id
                node_shape_id = f"node_{i}"
                node_item.setData(0, Qt.UserRole, node_shape_id)
                node_indices_for_tree.append(node_shape_id)

                # 保存节点 ref 到 shape_id 的映射
                self.node_id_map[node_ref] = node_shape_id

                # 存储节点信息，用于查询 (using node_shape_id as key)
                self.shape_to_info[node_shape_id] = {
                    'ref': node_ref, # Store original reference name
                    'coordinates': node_pos,
                    'connected_links': self.node_to_links.get(node_ref, [])
                }

                # 添加节点坐标信息
                x_item = QTreeWidgetItem(node_item)
                x_item.setText(0, f"X= {node_pos[0]}")
                y_item = QTreeWidgetItem(node_item)
                y_item.setText(0, f"Y= {node_pos[1]}")
                z_item = QTreeWidgetItem(node_item)
                z_item.setText(0, f"Z= {node_pos[2]}")

            # 设置节点根节点的索引列表
            nodes_root.setData(0, Qt.UserRole, node_indices_for_tree)

            # 设置每个section组的索引列表
            for section, section_item in section_groups.items():
                section_item.setData(0, Qt.UserRole, section_indices[section])

            # 将所有索引添加到根节点 (包括线段和节点 shape_ids)
            all_indices = []
            for indices in section_indices.values():
                all_indices.extend(indices) # Add segment indices (integers)
            all_indices.extend(node_indices_for_tree) # Add node shape_ids (strings)
            main_root.setData(0, Qt.UserRole, all_indices)

            # 关闭进度对话框
            progress.setValue(len(df))

            # 显示成功消息
            self.show_success_message(f"Excel数据解析完成! 提取了 {len(self.unique_nodes)} 个唯一节点。")

            # 展开第一层节点
            self.tree.expandItem(main_root)

            logger.info(f"解析完成，共 {len(self.segments)} 条线段和 {len(self.unique_nodes)} 个节点")

        except Exception as e:
            logger.error(f"创建树节点时出错: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.warning(self, "解析错误", f"构建树时出现错误: {str(e)}")

    def create_node_shapes(self):
        """创建代表节点的球体形状 (TopoDS_Shape)"""
        logger.info(f"创建节点形状，节点数量: {len(self.unique_nodes)}")
        self.node_shapes = []  # 重置节点 TopoDS_Shape 列表

        # 设置球体的半径
        radius = 40.0  # 球体的半径，可以根据需要调整

        for i, (node_ref, node_pos) in enumerate(self.unique_nodes.items()):
            try:
                # 创建球体
                center = gp_Pnt(*node_pos)
                sphere = BRepPrimAPI_MakeSphere(center, radius).Shape()
                if sphere.IsNull():
                    logger.warning(f"创建节点 {node_ref} 的球体失败")
                    continue
                self.node_shapes.append(sphere)
                # Note: We store the TopoDS_Shape here.
                # The AIS_Shape will be created in draw_segments.
            except Exception as e:
                logger.error(f"创建节点 {node_ref} 的形状时出错: {str(e)}")
                logger.error(traceback.format_exc())
                continue

        logger.info(f"节点 TopoDS_Shape 创建完成，成功创建: {len(self.node_shapes)}")


    def setup_interaction(self):
        """设置3D视图的交互功能"""
        try:
            # 注册鼠标点击事件回调
            # The callback signature might vary slightly depending on backend/version
            # Let's assume it returns a list of selected TopoDS_Shape objects
            self.viewer._display.register_select_callback(self.shape_selection_callback)

            # 设置状态栏初始信息
            self.status_bar.showMessage("点击3D形状以显示详细信息")
            logger.info("3D视图交互功能设置完成")
        except Exception as e:
            logger.error(f"设置交互功能时出错: {str(e)}")
            logger.error(traceback.format_exc())

    # --- CORRECTED shape_selection_callback ---
    def shape_selection_callback(self, shape_list, *args):
        """
        当用户在3D视图中点击形状时的回调函数 (Corrected Version)

        Parameters:
        -----------
        shape_list : list
            A list containing the selected TopoDS_Shape objects.
        *args : tuple
            Additional arguments like click coordinates (often (x, y)).
        """
        try:
            logger.debug(f"选择回调被触发，shape_list类型: {type(shape_list)}, 内容: {shape_list}")
            if args:
                logger.debug(f"附加参数: {args}")

            if not shape_list:
                logger.debug("没有选中任何形状 (shape_list is empty)")
                # Optionally clear selection/info here if nothing is clicked
                if self.highlighted_shapes:
                   self.highlight_shapes(self.highlighted_shapes, False)
                   self.highlighted_shapes = []
                   self.clear_info()
                return

            # We usually care about the first selected shape
            # The object in the list should be a TopoDS_Shape
            selected_shape = shape_list[0]
            if not isinstance(selected_shape, TopoDS_Shape) or selected_shape.IsNull():
                logger.warning(f"选择回调收到的不是有效的 TopoDS_Shape: {type(selected_shape)}")
                return

            logger.debug(f"选中的 TopoDS_Shape 类型: {self.get_shape_type_name(selected_shape)}")

            # --- Find the shape_id by comparing TopoDS_Shapes ---
            found_shape_id = None
            for shape_id, ais_obj in self.ais_shapes.items():
                try:
                    # Get the underlying TopoDS_Shape from the AIS_Shape
                    ais_topo_shape = ais_obj.Shape()
                    if not ais_topo_shape.IsNull():
                        # Use IsSame() for robust comparison
                        if ais_topo_shape.IsSame(selected_shape):
                            found_shape_id = shape_id
                            logger.debug(f"找到匹配的形状ID: {found_shape_id} (通过 IsSame 比较)")
                            break
                except Exception as e:
                    logger.warning(f"比较形状 {shape_id} 时出错: {e}")
                    continue

            if found_shape_id is None:
                logger.debug("未找到与选择形状匹配的ID")
                self.status_bar.showMessage("已选择形状，但无法在内部映射中找到它")
                 # Optionally clear selection/info here
                if self.highlighted_shapes:
                   self.highlight_shapes(self.highlighted_shapes, False)
                   self.highlighted_shapes = []
                   self.clear_info()
                return

            # --- Process the found shape ---

            # Clear previous highlight BEFORE applying new one
            if self.highlighted_shapes:
                # Avoid de-highlighting if clicking the same shape again
                if self.highlighted_shapes != [found_shape_id]:
                    self.highlight_shapes(self.highlighted_shapes, False)

            # Highlight the current selection
            self.highlight_shapes([found_shape_id], True)
            self.highlighted_shapes = [found_shape_id] # Store the *found* ID

            # Display info based on the type of ID found
            if isinstance(found_shape_id, str) and found_shape_id.startswith('node_'):
                self.display_node_info(found_shape_id)
            elif isinstance(found_shape_id, int): # Segments use integer indices
                self.display_link_info(found_shape_id)
            elif isinstance(found_shape_id, str): # Imported STEP/IGES shapes use string IDs
                logger.info(f"选中了导入的形状: {found_shape_id}")
                shape_info = self.shape_to_info.get(found_shape_id, {'type': 'Imported Shape'})
                info_text = f"选中的形状 ID: {found_shape_id}\n"
                info_text += f"类型: {shape_info.get('type', '未知')}"
                # Add more details if available in shape_to_info for imported shapes
                self.info_text.setText(info_text)
                self.status_bar.showMessage(f"已选择导入的形状: {found_shape_id}")
            else:
                 logger.warning(f"未知的 shape_id 类型: {type(found_shape_id)}")
                 self.clear_info()
                 self.status_bar.showMessage(f"选择了未知类型的形状: {found_shape_id}")


            # Find and select the corresponding item in the tree
            self.find_and_select_tree_item(found_shape_id)

        except Exception as e:
            logger.error(f"形状选择回调中出现错误: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"选择处理错误: {str(e)}")
    # --- End of CORRECTED shape_selection_callback ---


    def find_and_select_tree_item(self, shape_id):
        """根据形状ID查找并选择树中对应的项"""
        try:
            # Use an iterative approach to avoid deep recursion issues
            items_to_check = []
            root = self.tree.invisibleRootItem()
            for i in range(root.childCount()):
                items_to_check.append(root.child(i))

            found_item = None
            while items_to_check:
                current_item = items_to_check.pop(0) # Process like a queue
                item_data = current_item.data(0, Qt.UserRole)

                match = False
                if item_data == shape_id:
                    match = True
                elif isinstance(item_data, list) and shape_id in item_data:
                     # If data is a list, we need to find the specific child
                     # For now, let's just select the parent group item
                     match = True # Select the group if the ID is in its list

                if match:
                    found_item = current_item
                    break # Found it

                # Add children to the check list
                for i in range(current_item.childCount()):
                    items_to_check.append(current_item.child(i))

            if found_item:
                # Select and expand to the item
                self.tree.setCurrentItem(found_item)
                self.tree.scrollToItem(found_item, QTreeWidget.ScrollHint.PositionAtCenter) # Better scrolling
                # Ensure all parent items are expanded
                parent = found_item.parent()
                while parent and parent != self.tree.invisibleRootItem(): # Check parent is valid
                    self.tree.expandItem(parent)
                    parent = parent.parent()
                logger.debug(f"在树中找到并选择了项: {found_item.text(0)} for ID {shape_id}")
            else:
                logger.debug(f"未找到与形状ID {shape_id} 对应的树项")
        except Exception as e:
            logger.error(f"查找和选择树项时出错: {str(e)}")
            logger.error(traceback.format_exc())

    def display_node_info(self, node_shape_id):
        """显示节点信息 using node_shape_id ('node_i')"""
        try:
            if node_shape_id not in self.shape_to_info:
                logger.warning(f"节点信息中找不到节点 shape_id: {node_shape_id}")
                self.info_text.setText(f"找不到节点信息: {node_shape_id}")
                return

            node_info = self.shape_to_info[node_shape_id]
            node_ref = node_info.get('ref', '未知Ref') # Get original reference name
            coords = node_info.get('coordinates', (0,0,0))
            connected_links = node_info.get('connected_links', [])

            # 构建信息文本
            info_text = f"节点 Ref: {node_ref}\n"
            info_text += f"(内部 ID: {node_shape_id})\n"
            info_text += f"坐标: X={coords[0]:.2f}, Y={coords[1]:.2f}, Z={coords[2]:.2f}\n"
            info_text += f"连接的链接数: {len(connected_links)}\n\n"

            # 添加连接的链接信息
            if connected_links:
                info_text += "连接的链接:\n"
                for link_idx in connected_links[:10]:  # 最多显示前10个
                    if link_idx in self.link_data:
                        link = self.link_data[link_idx]
                        info_text += f"- {link['name']} (ID: {link_idx})\n"

                if len(connected_links) > 10:
                    info_text += f"... 以及其他 {len(connected_links) - 10} 个链接\n"

            # 更新信息显示
            self.info_text.setText(info_text)

            # 更新状态栏
            self.status_bar.showMessage(f"已选择节点: {node_ref}")

            # 在日志中输出更详细的信息
            logger.info(f"显示节点信息: {node_ref} (ID: {node_shape_id})")
            logger.debug(f"节点坐标: X={coords[0]}, Y={coords[1]}, Z={coords[2]}")
            logger.debug(f"连接的链接数: {len(connected_links)}")
        except Exception as e:
            logger.error(f"显示节点信息时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"显示节点信息时出错: {str(e)}")

    def display_link_info(self, link_idx):
        """显示链接信息 using link index"""
        try:
            if link_idx not in self.link_data:
                logger.warning(f"链接数据中找不到索引: {link_idx}")
                self.info_text.setText(f"找不到链接信息: {link_idx}")
                return

            link = self.link_data[link_idx]

            # 构建信息文本
            info_text = f"链接: {link['name']}\n"
            info_text += f"(内部 ID: {link_idx})\n\n"

            # 起始点信息
            origin = link['origin']
            info_text += f"起始点 Ref: {origin['ref']}\n"
            info_text += f"坐标: X={origin['coordinates'][0]:.2f}, Y={origin['coordinates'][1]:.2f}, Z={origin['coordinates'][2]:.2f}\n\n"

            # 终止点信息
            extremite = link['extremite']
            info_text += f"终止点 Ref: {extremite['ref']}\n"
            info_text += f"坐标: X={extremite['coordinates'][0]:.2f}, Y={extremite['coordinates'][1]:.2f}, Z={extremite['coordinates'][2]:.2f}\n\n"

            # 其他信息
            info_text += f"长度: {link.get('length', 'N/A')}\n" # Use .get for safety
            info_text += f"密度: {link.get('density', 'N/A')}\n"
            info_text += f"安全等级: {link.get('safety', 'N/A')}\n"
            info_text += f"路由: {link.get('route', 'N/A')}\n"
            info_text += f"动作号: {link.get('action_number', 'N/A')}\n"
            info_text += f"截面: {link.get('section', 'N/A')}\n"

            # 更新信息显示
            self.info_text.setText(info_text)

            # 更新状态栏
            self.status_bar.showMessage(f"已选择链接: {link['name']}")

            # 在日志中输出更详细的信息
            logger.info(f"显示链接信息: {link['name']} (ID: {link_idx})")
            logger.debug(f"起点: {origin['ref']} 坐标: ({origin['coordinates'][0]}, {origin['coordinates'][1]}, {origin['coordinates'][2]})")
            logger.debug(f"终点: {extremite['ref']} 坐标: ({extremite['coordinates'][0]}, {extremite['coordinates'][1]}, {extremite['coordinates'][2]})")
        except Exception as e:
            logger.error(f"显示链接信息时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"显示链接信息时出错: {str(e)}")

    def clear_info(self):
        """清除显示的信息"""
        self.info_text.clear()
        self.status_bar.showMessage("点击3D形状以显示详细信息")

    def draw_segments(self):
        """Draw all segments (cylinders) and nodes (spheres) in the 3D viewer."""
        try:
            logger.info("开始绘制线段和节点...")
            # Don't erase if we are adding to existing imported shapes (future feature?)
            # For now, assume we clear when drawing segments derived from Excel
            self.context.EraseAll(False) # Erase, but don't redraw yet
            self.segment_shapes = []  # Reset TopoDS_Shape list for segments
            self.ais_shapes = {}  # Clear AIS shape dictionary
            self.highlighted_shapes = []  # Clear highlight list

            # Create node TopoDS_Shapes first (if not already done)
            if not self.node_shapes and self.unique_nodes:
                 self.create_node_shapes()

            # Determine total shapes for progress
            total_shapes_to_draw = len(self.segments) + len(self.node_shapes)
            if total_shapes_to_draw == 0:
                 logger.info("没有线段或节点可绘制")
                 self.viewer._display.Repaint()
                 return

            progress = None
            if self.first_draw:
                progress = QProgressDialog("正在绘制线段和节点...", "取消", 0, total_shapes_to_draw, self)
                progress.setWindowModality(Qt.WindowModal)
                progress.setMinimumDuration(500)
                progress.show()
                QCoreApplication.processEvents()

            shape_counter = 0

            # Create and display nodes (Spheres)
            # Map node ref to its index in self.node_shapes
            node_ref_list = list(self.unique_nodes.keys())
            for i, sphere_shape in enumerate(self.node_shapes):
                if progress and progress.wasCanceled(): break
                try:
                    # Get the corresponding node_shape_id ('node_i')
                    node_ref = node_ref_list[i] # Get ref using the index
                    node_shape_id = self.node_id_map.get(node_ref) # Look up 'node_i' id
                    if node_shape_id is None:
                        logger.warning(f"无法找到节点 ref '{node_ref}' 的 shape_id 映射")
                        node_shape_id = f"node_{i}" # Fallback ID

                    # Create AIS_Shape for the node
                    ais_sphere = AIS_Shape(sphere_shape)
                    color = Quantity_Color(Quantity_NOC_RED)
                    self.context.SetColor(ais_sphere, color, False)
                    self.context.Display(ais_sphere, False) # Display without immediate update

                    # Store AIS object with its shape_id
                    self.ais_shapes[node_shape_id] = ais_sphere

                    shape_counter += 1
                    if progress: progress.setValue(shape_counter)
                    QCoreApplication.processEvents()

                except Exception as e:
                    logger.error(f"显示节点 {i} (Ref: {node_ref}) 时出错: {str(e)}")
                    continue

            if progress and progress.wasCanceled():
                 logger.info("用户取消了绘制操作")
                 self.context.UpdateCurrentViewer() # Update viewer with what was drawn
                 return

            # Create and display segments (Cylinders)
            self.segment_shapes = [] # Rebuild segment TopoDS list
            for i, (start, end) in enumerate(self.segments):
                if progress and progress.wasCanceled(): break
                try:
                    start_pnt = gp_Pnt(*start)
                    end_pnt = gp_Pnt(*end)
                    height = start_pnt.Distance(end_pnt)
                    # Handle zero-length segments gracefully
                    if height < 1e-6: # Use a small tolerance
                         logger.warning(f"线段 {i} 长度接近零，跳过绘制圆柱体")
                         # Optionally draw a small sphere or point instead
                         continue

                    radius = 30.0 # Cylinder radius

                    # Calculate direction, handle potential zero vector if points coincide (already checked by height)
                    vec = gp_Vec(start_pnt, end_pnt)
                    if vec.Magnitude() < 1e-6: continue # Double check
                    direction = gp_Dir(vec)
                    axis = gp_Ax2(start_pnt, direction)

                    # Create cylinder
                    cylinder = BRepPrimAPI_MakeCylinder(axis, radius, height).Shape()
                    if cylinder.IsNull():
                        logger.warning(f"创建线段 {i} 的圆柱体失败")
                        continue

                    self.segment_shapes.append(cylinder) # Store TopoDS_Shape

                    # Create AIS_Shape for the segment
                    ais_cylinder = AIS_Shape(cylinder)
                    color = Quantity_Color(Quantity_NOC_BLUE)
                    self.context.SetColor(ais_cylinder, color, False)
                    self.context.Display(ais_cylinder, False) # Display without immediate update

                    # Store AIS object using segment index as shape_id
                    self.ais_shapes[i] = ais_cylinder

                    shape_counter += 1
                    if progress: progress.setValue(shape_counter)
                    QCoreApplication.processEvents()

                except Exception as e:
                    logger.error(f"创建或显示线段 {i} 时出错: {str(e)}")
                    logger.error(traceback.format_exc())
                    continue

            # Close progress dialog and update viewer once
            if progress:
                progress.setValue(total_shapes_to_draw)

            self.viewer._display.FitAll()
            self.context.UpdateCurrentViewer() # Update the viewer to show all changes
            # self.viewer._display.Repaint() # Redundant if UpdateCurrentViewer is called

            self.first_draw = False # Mark as drawn

            logger.info(f"完成绘制，线段数量: {len(self.segment_shapes)}，节点数量: {len(self.node_shapes)}")
        except Exception as e:
            logger.error(f"绘制线段和节点时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"绘制图形时出错: {str(e)}")
            if progress: progress.close()
            # Ensure viewer is updated even on error
            self.context.UpdateCurrentViewer()


    def draw_imported_shapes(self, show_progress=True):
        """显示导入的STEP/IGES形状，使用AIS_Shape进行管理。"""
        try:
            logger.info(f"开始显示导入的形状，形状数量: {len(self.step_shapes)}")
            self.context.EraseAll(False)  # 清除现有显示, no update yet
            self.ais_shapes = {}  # 清空AIS形状字典
            self.highlighted_shapes = []  # 清空高亮列表

            if not self.step_shapes:
                 logger.info("没有导入的形状可显示")
                 self.context.UpdateCurrentViewer()
                 return

            progress = None
            if show_progress:
                total_shapes = len(self.step_shapes)
                progress = QProgressDialog("正在显示导入的形状...", "取消", 0, total_shapes, self)
                progress.setWindowModality(Qt.WindowModal)
                progress.setMinimumDuration(500)
                progress.show()
                QCoreApplication.processEvents()

            shapes_displayed = 0
            for i, (shape_id, topo_shape) in enumerate(self.step_shapes.items()):
                if progress and progress.wasCanceled(): break
                try:
                    if topo_shape.IsNull():
                        logger.warning(f"跳过空的导入形状: {shape_id}")
                        continue

                    # Create AIS_Shape for the imported shape
                    ais_imported = AIS_Shape(topo_shape)

                    # Set default color (e.g., gray or white for imported base models)
                    color = Quantity_Color(0.8, 0.8, 0.8, Quantity_TOC_RGB) # Light Gray
                    self.context.SetColor(ais_imported, color, False)
                    self.context.Display(ais_imported, False) # Display without immediate update

                    # Store AIS object
                    self.ais_shapes[shape_id] = ais_imported
                    shapes_displayed += 1

                    if progress: progress.setValue(i + 1)
                    QCoreApplication.processEvents()

                except Exception as e:
                    logger.error(f"无法显示形状 {shape_id}: {e}")
                    logger.error(traceback.format_exc())
                    continue

            # Close progress dialog and update viewer
            if progress:
                progress.setValue(len(self.step_shapes))

            self.viewer._display.FitAll()
            self.context.UpdateCurrentViewer() # Update the viewer

            logger.info(f"完成显示，成功显示 {shapes_displayed}/{len(self.step_shapes)} 个形状")
        except Exception as e:
            logger.error(f"显示导入形状时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"显示形状时出错: {str(e)}")
            if progress: progress.close()
            self.context.UpdateCurrentViewer() # Ensure update on error


    def highlight_shapes(self, shape_ids, highlight=True):
        """高亮显示或取消高亮指定的形状 (Optimized)."""
        try:
            if not isinstance(shape_ids, list):
                shape_ids = [shape_ids]

            logger.debug(f"{'高亮' if highlight else '取消高亮'} {len(shape_ids)} 个形状: {shape_ids[:5]}...") # Log first few

            needs_update = False
            for shape_id in shape_ids:
                if shape_id in self.ais_shapes:
                    ais_obj = self.ais_shapes[shape_id]
                    try:
                        if highlight:
                            # Use context's highlight/selection methods for better visual distinction
                            # self.context.ClearSelected(False) # Clear previous interactive selections if needed
                            # self.context.AddOrRemoveSelected(ais_obj, True) # Select the object
                            # Use color for now as selection might interfere with callbacks
                            if isinstance(shape_id, str) and shape_id.startswith('node_'):
                                                                               color = Quantity_Color(Quantity_NOC_GREEN)
                            else: # Segments or imported shapes
                                color = Quantity_Color(Quantity_NOC_YELLOW)
                            self.context.SetColor(ais_obj, color, False)
                        else:
                            # Restore default color
                            # self.context.RemoveSelected(ais_obj, False) # Deselect if using selection
                            if isinstance(shape_id, str) and shape_id.startswith('node_'):
                                color = Quantity_Color(Quantity_NOC_RED) # Default node color
                            elif isinstance(shape_id, int): # Segment
                                color = Quantity_Color(Quantity_NOC_BLUE) # Default segment color
                            else: # Imported shape
                                color = Quantity_Color(0.8, 0.8, 0.8, Quantity_TOC_RGB) # Default imported color
                            self.context.SetColor(ais_obj, color, False)

                        needs_update = True # Mark that viewer needs update
                    except Exception as e:
                        logger.error(f"设置形状 {shape_id} 的高亮/颜色时出错: {e}")
                        continue
                else:
                    logger.warning(f"尝试高亮/取消高亮时找不到 shape_id: {shape_id}")


            # Update the viewer only once after processing all IDs if changes were made
            if needs_update:
                try:
                    # Update only the presentation without recomputing structure
                    self.context.UpdateCurrentViewer()
                    # For color changes, sometimes redisplay might be needed if Update doesn't work
                    # self.context.Redisplay(ais_obj, True) # Do this inside the loop if UpdateCurrentViewer isn't enough
                except Exception as e:
                    logger.error(f"更新查看器以应用高亮时出错: {e}")

        except Exception as e:
            logger.error(f"高亮形状时出错: {str(e)}")
            logger.error(traceback.format_exc())


    def on_tree_item_clicked(self, item, column):
        """Handle tree item clicks to highlight shapes."""
        try:
            shape_data = item.data(0, Qt.UserRole)

            logger.debug(f"点击了树项: {item.text(0)}, 数据: {shape_data}")

            if shape_data is None:
                logger.debug("树项没有关联的形状数据")
                # Option: Clear selection if clicking an item with no data
                if self.highlighted_shapes:
                    self.highlight_shapes(self.highlighted_shapes, False)
                    self.highlighted_shapes = []
                    self.clear_info()
                return

            ids_to_highlight = []
            if isinstance(shape_data, list):
                ids_to_highlight.extend(shape_data)
            else: # Single ID (int for segment, string for node/imported)
                ids_to_highlight.append(shape_data)

            # Filter out invalid IDs before proceeding
            valid_ids_to_highlight = [sid for sid in ids_to_highlight if sid in self.ais_shapes]
            if not valid_ids_to_highlight:
                 logger.debug(f"树项关联的ID {ids_to_highlight} 在ais_shapes中均未找到")
                 # Clear previous selection if clicking something non-highlightable
                 if self.highlighted_shapes:
                    self.highlight_shapes(self.highlighted_shapes, False)
                    self.highlighted_shapes = []
                    self.clear_info()
                 return


            # --- Handle highlighting ---
            # If clicking the same item/group, deselect
            if self.selected_item == item:
                 logger.debug("再次点击同一项目，取消高亮")
                 self.highlight_shapes(self.highlighted_shapes, False)
                 self.highlighted_shapes = []
                 self.selected_item = None
                 self.clear_info()
                 self.tree.clearSelection() # Visually deselect in tree
            else:
                # Clear previous highlight
                if self.highlighted_shapes:
                    self.highlight_shapes(self.highlighted_shapes, False)

                # Highlight the new selection
                self.highlight_shapes(valid_ids_to_highlight, True)
                self.highlighted_shapes = valid_ids_to_highlight
                self.selected_item = item

                # --- Display Info for the first highlighted item ---
                if valid_ids_to_highlight:
                    first_id = valid_ids_to_highlight[0]
                    if isinstance(first_id, str) and first_id.startswith('node_'):
                        self.display_node_info(first_id)
                    elif isinstance(first_id, int):
                        self.display_link_info(first_id)
                    elif isinstance(first_id, str): # Imported shape
                         logger.info(f"选中了导入的形状 (从树): {first_id}")
                         shape_info = self.shape_to_info.get(first_id, {'type': 'Imported Shape'})
                         info_text = f"选中的形状 ID: {first_id}\n"
                         info_text += f"类型: {shape_info.get('type', '未知')}\n"
                         if len(valid_ids_to_highlight) > 1:
                             info_text += f"\n(以及其他 {len(valid_ids_to_highlight)-1} 个形状)"
                         self.info_text.setText(info_text)
                         self.status_bar.showMessage(f"已选择 {len(valid_ids_to_highlight)} 个形状")
                    else:
                         self.clear_info()
                else:
                     self.clear_info()


        except Exception as e:
            logger.error(f"处理树项点击时出错: {str(e)}")
            logger.error(traceback.format_exc())
            self.status_bar.showMessage(f"处理选择时出错: {str(e)}")


    def closeEvent(self, event):
        """窗口关闭时清理资源"""
        try:
            logger.info("正在关闭窗口，清理资源...")

            # 1. 清除所有显示的图形 from context
            if hasattr(self, 'context') and self.context:
                try:
                    self.context.EraseAll(True)  # Erase and update immediately
                    logger.debug("已清除所有显示的图形 (Context)")
                except Exception as e:
                    logger.error(f"清除图形时出错: {e}")

            # 2. 关闭并释放 viewer 对象
            if hasattr(self, 'viewer') and self.viewer:
                try:
                    # Disconnect callbacks explicitly if possible
                    # (Might not be necessary if viewer handles it on close)
                    # self.viewer._display.unregister_select_callback(self.shape_selection_callback) # Example

                    self.viewer.close()  # Close viewer widget
                    logger.debug("已关闭viewer")
                    self.viewer.deleteLater()  # Schedule for deletion
                    logger.debug("已安排删除viewer")
                    self.viewer = None
                    self.context = None # Clear context reference too
                except Exception as e:
                    logger.error(f"关闭viewer时出错: {e}")

            # 3. 清理计时器
            if hasattr(self, 'timer') and self.timer:
                try:
                    self.timer.stop()
                    logger.debug("已停止计时器")
                    if hasattr(self, 'timer_connected') and self.timer_connected:
                        try:
                            # Safely disconnect
                            self.timer.timeout.disconnect(self.close_message_box)
                            logger.debug("已断开计时器连接")
                        except (TypeError, RuntimeError) as e:
                            logger.warning(f"断开计时器连接时出错 (可能已断开): {e}")
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

    def export_file(self):
        """根据选择的格式导出文件"""
        file_format = self.export_format_combo.currentText()
        logger.info(f"用户选择导出文件格式: {file_format}")

        # Determine what to export
        shapes_to_export = []
        export_description = ""

        if self.main_shape and not self.main_shape.IsNull():
            shapes_to_export.append(self.main_shape)
            export_description = "主导入模型"
            logger.info("准备导出主导入模型...")
        elif self.step_shapes:
             # Export all individual imported shapes as a compound
             shapes_to_export.extend(self.step_shapes.values())
             export_description = f"{len(shapes_to_export)} 个导入的形状"
             logger.info(f"准备导出 {len(shapes_to_export)} 个导入的形状...")
        elif self.segment_shapes or self.node_shapes:
             # Export generated segments and nodes
             shapes_to_export.extend(self.segment_shapes)
             shapes_to_export.extend(self.node_shapes)
             export_description = f"{len(self.segment_shapes)} 条线段和 {len(self.node_shapes)} 个节点"
             logger.info(f"准备导出 {len(self.segment_shapes)} 条线段和 {len(self.node_shapes)} 个节点...")
        else:
            logger.warning("没有可导出的形状")
            QMessageBox.information(self, "无内容", "当前没有可导出的3D模型。请先导入文件或加载Excel数据。")
            return

        # Filter out null shapes just in case
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
                Interface_Static_SetCVal("write.step.schema", "AP203") # Or AP214

                progress.setValue(30)
                progress.setLabelText("正在处理形状...")
                QCoreApplication.processEvents()

                # Transfer shapes - if it's the single main_shape, transfer directly.
                # If multiple shapes, create a compound.
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
                             continue # Skip problematic shape

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
                # Set IGES parameters if needed
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
                    # IGES Write returns boolean, not status code like STEP
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
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 IGES 文件", "", "IGES 文件 (*.iges *.igs)")
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
            # --- Clear existing data ---
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
            self.step_shapes = {} # Clear previously imported shapes
            self.ais_shapes = {}  # Clear AIS objects
            self.highlighted_shapes = []
            self.main_shape = None
            self.context.EraseAll(True) # Erase and update viewer
            logger.info("现有数据清除完毕")
            progress.setValue(10)
            QCoreApplication.processEvents()

            # --- Read File ---
            logger.info(f"使用 {file_format} reader 读取文件...")
            if file_format == "STEP":
                reader = STEPControl_Reader()
            elif file_format == "IGES":
                reader = IGESControl_Reader()
            else:
                raise ValueError(f"不支持的文件格式: {file_format}")

            read_status = reader.ReadFile(file_path)
            logger.info(f"ReadFile status: {read_status}")
            progress.setValue(40)
            QCoreApplication.processEvents()
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")

            if read_status != IFSelect_RetDone:
                # Try to get more error info if possible (might not always work)
                fail_messages = reader.FailMessage()
                error_msg = f"读取{file_format}文件失败 (状态: {read_status})"
                if fail_messages:
                     error_msg += f"\n错误信息:\n{fail_messages}"
                     logger.error(f"{error_msg}")
                else:
                    logger.error(error_msg)
                raise RuntimeError(error_msg)


            # --- Transfer Roots ---
            logger.info("转换文件根...")
            # Check number of roots before transferring
            num_roots = reader.NbRootsForTransfer()
            logger.info(f"文件中根的数量: {num_roots}")
            if num_roots == 0:
                 raise RuntimeError(f"{file_format} 文件中没有找到可转换的根。文件可能为空或格式无效。")

            transfer_ok = reader.TransferRoots()
            logger.info(f"TransferRoots status: {transfer_ok}")
            progress.setValue(60)
            QCoreApplication.processEvents()
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")
            if not transfer_ok:
                raise RuntimeError(f"转换{file_format}文件根失败。")

            # --- Get Shape ---
            logger.info("获取主形状...")
            # Check number of resulting shapes
            num_shapes = reader.NbShapes()
            logger.info(f"转换后的形状数量: {num_shapes}")
            if num_shapes == 0:
                raise RuntimeError(f"转换{file_format}文件后未生成任何形状。")
            elif num_shapes == 1:
                logger.info("获取单个主形状...")
                shape = reader.Shape(1) # Get the first shape
            else:
                # If multiple shapes result, create a compound
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

            self.main_shape = shape # Store the main (possibly compound) shape
            logger.info(f"主形状类型: {self.get_shape_type_name(self.main_shape)}")
            progress.setValue(70)
            QCoreApplication.processEvents()
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")

            # --- Analyze and Build Tree ---
            logger.info("分析形状并构建树...")
            file_basename = os.path.basename(file_path)
            root = QTreeWidgetItem(self.tree)
            root.setText(0, f"{file_format} Model: {file_basename}")

            # Pass progress to allow cancellation during analysis
            self.analyze_shape_and_build_tree(self.main_shape, root, progress)
            if progress.wasCanceled(): raise InterruptedError("用户取消导入")

            self.tree.expandItem(root)
            progress.setValue(90)
            QCoreApplication.processEvents()

            # --- Draw Imported Shapes ---
            logger.info("绘制导入的形状...")
            self.first_draw = True # Ensure progress bar shows for drawing
            self.draw_imported_shapes(show_progress=True) # Pass progress? No, draw has its own
            if progress.wasCanceled(): raise InterruptedError("用户取消导入") # Check again after drawing

            progress.setValue(100)

            total_elements = len(self.step_shapes) # Count individual shapes stored
            logger.info(f"{file_format}文件导入成功，包含 {total_elements} 个子元素")
            self.show_success_message(f"已成功导入 {file_format} 文件 ({file_basename})，包含 {total_elements} 个元素。")

        except InterruptedError:
             logger.warning(f"用户取消了 {file_format} 导入")
             progress.close()
             # Clean up potentially partially loaded state
             self.tree.clear()
             self.context.EraseAll(True)
             self.status_bar.showMessage(f"{file_format} 导入已取消")
        except Exception as e:
            progress.close()
            logger.error(f"导入{file_format}文件时发生异常: {e}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "导入错误", f"导入 {file_path} 时出现异常:\n\n{str(e)}\n\n请查看日志获取详细信息。")
            # Clean up
            self.tree.clear()
            self.context.EraseAll(True)


    def analyze_shape_and_build_tree(self, shape, parent_item, progress=None):
        """分析形状的层次结构并构建树视图 (Builds tree, populates self.step_shapes)."""
        try:
            logger.info("开始分析形状...")
            if progress: progress.setLabelText("分析形状结构...")

            # Clear previous imported shapes before populating
            self.step_shapes = {}
            all_shape_ids_in_tree = [] # Track IDs added to the tree structure

            # --- Use TopExp_Explorer to iterate through sub-shapes ---
            # We primarily care about Solids, Shells, Faces, Edges for visualization/selection
            shape_types_to_explore = {
                TopAbs_SOLID: "实体 (Solid)",
                TopAbs_SHELL: "壳 (Shell)",
                TopAbs_FACE: "面 (Face)",
                # TopAbs_WIRE: "线框 (Wire)", # Often part of faces, might add noise
                TopAbs_EDGE: "边 (Edge)"
            }
            type_nodes = {} # Store top-level type nodes in the tree

            # Estimate total work for progress update
            # Counting shapes beforehand can be slow for complex models.
            # We'll update progress based on iteration.

            processed_count = 0
            unique_subshapes = TopTools_IndexedMapOfShape() # Use map to avoid duplicates

            # First pass: Collect unique subshapes of relevant types
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


            # Second pass: Iterate unique shapes, create tree items, store shapes
            logger.debug("第二遍: 创建树节点并存储形状...")
            shape_counter = 0
            for i in range(1, total_unique + 1):
                if progress and progress.wasCanceled(): raise InterruptedError("用户取消分析")

                current_shape = unique_subshapes.FindKey(i)
                if current_shape.IsNull(): continue

                shape_type = current_shape.ShapeType()
                type_name = self.get_shape_type_name(current_shape) # Get user-friendly name
                base_type_name = shape_types_to_explore.get(shape_type, "其他") # Get category name

                # Get or create the category node in the tree
                if base_type_name not in type_nodes:
                    type_node = QTreeWidgetItem(parent_item)
                    type_node.setText(0, base_type_name)
                    type_nodes[base_type_name] = {"node": type_node, "ids": []}
                else:
                    type_node = type_nodes[base_type_name]["node"]

                # Create item for the individual shape
                shape_item = QTreeWidgetItem(type_node)
                 # Generate a unique ID: type_index_hash
                shape_id = f"{base_type_name.split(' ')[0].lower()}_{shape_counter}_{current_shape.HashCode(1000000)}"
                shape_item.setText(0, f"{type_name} ID: {shape_id}")
                shape_item.setData(0, Qt.UserRole, shape_id) # Store ID in item

                # Store the actual TopoDS_Shape in our dictionary
                self.step_shapes[shape_id] = current_shape
                 # Store basic info (can be expanded later)
                self.shape_to_info[shape_id] = {'type': type_name}

                # Add ID to lists for group highlighting
                type_nodes[base_type_name]["ids"].append(shape_id)
                all_shape_ids_in_tree.append(shape_id)
                shape_counter += 1


                processed_count += 1
                if progress and total_unique > 0:
                     prog_val = 70 + int(20 * processed_count / total_unique) # Progress from 70% to 90%
                     progress.setValue(prog_val)
                     QCoreApplication.processEvents()


            # Update count on type nodes and store ID lists in them
            logger.debug("更新树节点计数和数据...")
            for base_type_name, data in type_nodes.items():
                 node = data["node"]
                 ids = data["ids"]
                 node.setText(0, f"{base_type_name} ({len(ids)})")
                 node.setData(0, Qt.UserRole, ids)

            # Store all collected IDs in the root item as well
            parent_item.setData(0, Qt.UserRole, all_shape_ids_in_tree)

            # Store the main shape itself if it wasn't broken down (e.g., a simple solid)
            # We already stored self.main_shape during import

            logger.info(f"形状分析完成，存储了 {len(self.step_shapes)} 个形状，创建了 {len(all_shape_ids_in_tree)} 个树条目")

        except InterruptedError:
             logger.warning("形状分析被用户取消")
             raise # Re-raise to be caught by importer
        except Exception as e:
            logger.error(f"分析形状和构建树时出错: {e}")
            logger.error(traceback.format_exc())
            # Don't raise here, allow import to potentially finish but show warning
            QMessageBox.warning(self, "分析警告", f"分析形状结构时出现错误:\n{e}\n导入结果可能不完整。")


    def get_shape_type_name(self, shape):
        """获取形状类型的用户友好名称"""
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
            # logger.warning(f"获取形状类型名称时出错: {e}") # Can be noisy
            return "未知形状"


    def show_success_message(self, message):
        """显示成功消息，并在 1.5 秒后自动关闭"""
        try:
            # Close existing message box immediately if present
            if hasattr(self, 'msg_box') and self.msg_box and self.msg_box.isVisible():
                self.msg_box.close()
                # Ensure timer is stopped if it was running for the old box
                if hasattr(self, 'timer') and self.timer.isActive():
                    self.timer.stop()
                    if hasattr(self, 'timer_connected') and self.timer_connected:
                         try: self.timer.timeout.disconnect(self.close_message_box)
                         except: pass # Ignore errors
                         self.timer_connected = False

            # Create and show new message box
            self.msg_box = QMessageBox(self)
            self.msg_box.setWindowTitle("成功")
            self.msg_box.setText(message)
            self.msg_box.setIcon(QMessageBox.Information)
            # self.msg_box.setStandardButtons(QMessageBox.Ok) # No buttons needed for auto-close
            self.msg_box.setStandardButtons(QMessageBox.NoButton) # Hide OK button
            # Ensure it's modeless so the main window remains interactive
            self.msg_box.setWindowModality(Qt.NonModal)
            self.msg_box.show()

            # Set timer to close it
            if hasattr(self, 'timer'):
                # Disconnect previous connection just in case
                if hasattr(self, 'timer_connected') and self.timer_connected:
                    try: self.timer.timeout.disconnect(self.close_message_box)
                    except (TypeError, RuntimeError): pass
                # Connect and start timer
                self.timer.timeout.connect(self.close_message_box)
                self.timer_connected = True
                self.timer.start(1500)  # 1500 ms = 1.5 seconds

            logger.info(f"显示成功消息: {message}")
        except Exception as e:
            logger.error(f"显示成功消息时出错: {e}")
            logger.error(traceback.format_exc())

    def close_message_box(self):
        """Slot to close the success message box."""
        try:
            if hasattr(self, 'msg_box') and self.msg_box and self.msg_box.isVisible():
                self.msg_box.close()
                logger.debug("自动关闭成功消息框")

            # Disconnect timer signal
            if hasattr(self, 'timer_connected') and self.timer_connected:
                try:
                    self.timer.timeout.disconnect(self.close_message_box)
                except (TypeError, RuntimeError): pass # Ignore errors if already disconnected
                self.timer_connected = False
        except Exception as e:
            logger.error(f"关闭消息框时出错: {e}")
            logger.error(traceback.format_exc())


def main(xlsx_file=None):
    # 设置日志系统
    log_file = setup_logging()

    # Create QApplication instance earlier
    app = QApplication.instance() # Check if already exists
    if not app: # Create if does not exist
        app = QApplication(sys.argv)

    logger.info(f"应用程序启动，日志文件: {log_file}")
    logger.info(f"Python版本: {sys.version}")
    logger.info(f"Qt版本: {QT_VERSION_STR}")
    logger.info(f"OCC Backend: qt-pyqt5") # Confirm backend

    # 显示命令行参数
    logger.info(f"命令行参数: {sys.argv}")
    if xlsx_file:
        logger.info(f"指定的Excel文件: {xlsx_file}")
    else:
        logger.info("未指定Excel文件，将启动空窗口。")


    df = None # Initialize df to None
    if xlsx_file:
        if not os.path.exists(xlsx_file):
            logger.error(f"指定的Excel文件不存在: {xlsx_file}")
            print(f"错误: 文件未找到 '{xlsx_file}'")
            # Optionally show a GUI error message here before creating the window
            error_msg = QMessageBox()
            error_msg.setIcon(QMessageBox.Critical)
            error_msg.setWindowTitle("文件错误")
            error_msg.setText(f"无法找到指定的Excel文件:\n{xlsx_file}\n\n应用程序将以空状态启动。")
            error_msg.exec_()
            xlsx_file = None # Proceed without the file
        else:
            try:
                logger.info(f"正在读取Excel文件: {xlsx_file}")
                # Try specifying engine if default fails on some xlsx files
                try:
                    df = pd.read_excel(xlsx_file, engine='openpyxl')
                except ImportError:
                    logger.warning("openpyxl 未安装，尝试默认引擎")
                    df = pd.read_excel(xlsx_file)

                logger.info(f"Excel读取成功，行数: {len(df)}, 列数: {len(df.columns)}")
                logger.debug(f"列名: {df.columns.tolist()}")
            except Exception as e:
                logger.error(f"读取Excel文件时出错: {str(e)}")
                logger.error(traceback.format_exc())
                # Show GUI error message
                error_msg = QMessageBox()
                error_msg.setIcon(QMessageBox.Critical)
                error_msg.setWindowTitle("Excel 读取错误")
                error_msg.setText(f"无法读取Excel文件:\n{xlsx_file}\n\n错误: {str(e)}\n\n应用程序将以空状态启动。")
                error_msg.exec_()
                df = None # Ensure df is None if reading failed


    # Create the main window (pass df which might be None)
    try:
        window = MainWindow(df)

        # Set window size and center
        window.resize(1200, 900) # Slightly larger default size
        try:
             screen_geometry = QDesktopWidget().availableGeometry() # Use availableGeometry
             x = (screen_geometry.width() - window.width()) // 2
             y = (screen_geometry.height() - window.height()) // 2
             window.move(x, y)
             logger.info(f"窗口大小设置为 {window.width()}x{window.height()}，居中显示在位置 ({x}, {y})")
        except Exception as e:
             logger.warning(f"无法获取屏幕几何信息或居中窗口: {e}")


        logger.info("显示主窗口")
        window.show()

        # Start the application event loop
        exit_code = app.exec_()
        logger.info(f"应用程序退出，退出代码: {exit_code}")
        return exit_code
    except Exception as e:
        logger.critical(f"创建或显示主窗口时发生未捕获异常: {e}")
        logger.critical(traceback.format_exc())
        # Show final critical error message
        error_msg = QMessageBox()
        error_msg.setIcon(QMessageBox.Critical)
        error_msg.setWindowTitle("应用程序错误")
        error_msg.setText(f"应用程序遇到严重错误并需要关闭:\n\n{str(e)}\n\n请查看日志文件 '{log_file}' 获取详细信息。")
        error_msg.exec_()
        return 1


if __name__ == "__main__":
    # Use argparse to parse command line arguments
    parser = argparse.ArgumentParser(description="基于公共数据源的航电系统布线架构与集成设计技术研究系统")
    parser.add_argument("xlsx_file", type=str, nargs='?', default=None,
                        help="Path to the Excel file containing wiring data.")
    parser.add_argument("--debug", action="store_true",
                        help="Enable detailed debug logging to console and file.")
    args = parser.parse_args()

    # Configure logging level based on debug flag BEFORE setting up handlers
    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.getLogger().setLevel(log_level) # Set root logger level

    if args.debug:
        print("--- Debug Mode Enabled ---")


    try:
        # Call the main function and exit with its return code
        exit_status = main(args.xlsx_file)
        sys.exit(exit_status)
    except Exception as e:
        # Catch any unexpected exceptions during startup or shutdown
        print(f"发生未捕获的顶层异常: {e}")
        traceback.print_exc()
        # Try to log it if logger is still available
        try:
             logger.critical(f"发生未捕获的顶层异常: {e}", exc_info=True)
        except:
             pass # Logger might not be initialized
        sys.exit(1)