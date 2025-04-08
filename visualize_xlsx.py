import sys
import argparse
from OCC.Core.Quantity import Quantity_Color
from OCC.Core._Quantity import Quantity_TOC_RGB
from OCC.Display.backend import load_backend
from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_Reader, STEPControl_AsIs
from OCC.Core.IGESControl import IGESControl_Reader, IGESControl_Writer
from OCC.Core.Interface import Interface_Static_SetCVal
from OCC.Core.IFSelect import IFSelect_RetDone
from OCC.Core.BRep import BRep_Tool
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
    QGroupBox
)
from OCC.Core.BRepPrimAPI import BRepPrimAPI_MakeCylinder, BRepPrimAPI_MakeSphere
from OCC.Core.gp import gp_Pnt, gp_Dir, gp_Ax2
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
from OCC.Core.BRep import BRep_Tool, BRep_Builder
from OCC.Core.AIS import AIS_Shape
from OCC.Core.Quantity import Quantity_Color, Quantity_NOC_BLUE, Quantity_NOC_YELLOW, Quantity_NOC_RED, Quantity_NOC_GREEN


class MainWindow(QWidget):
    def __init__(self, df=None):
        super().__init__()
        self.setWindowTitle("航电布线可视化系统")
        
        # 初始化消息框和计时器
        self.msg_box = None
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        
        # 标记计时器是否已连接
        self.timer_connected = False
        
        # 标记是否首次绘制
        self.first_draw = True
        
        # 高亮和颜色管理
        self.ais_shapes = {}  # 存储所有AIS对象，用于颜色管理
        self.highlighted_shapes = []  # 当前高亮的形状IDs

        # 节点相关数据结构
        self.unique_nodes = {}  # 存储唯一节点信息，使用ref作为键，(x, y, z)作为值
        self.node_shapes = []  # 存储节点的形状对象

        # 设置字体
        font = QFont()
        font.setPointSize(10)
        self.setFont(font)

        # 主布局
        main_layout = QHBoxLayout()

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

        main_layout.addLayout(left_layout, 1)

        # 3D 视图
        self.viewer = qtViewer3d(self)
        bg_color = Quantity_Color(0.2, 0.2, 0.2, Quantity_TOC_RGB)
        self.viewer._display.View.SetBackgroundColor(bg_color)
        main_layout.addWidget(self.viewer, 2)

        self.setLayout(main_layout)

        # 存储线段形状和原始颜色
        self.segment_shapes = []
        self.segments = []
        
        # 保存导入的模型
        self.step_shapes = {}  # 形状ID到形状对象的映射
        self.main_shape = None
        
        # 树项点击事件
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        self.selected_item = None

        # 如果有数据，则解析并显示
        if df is not None:
            self.parse_df_and_populate_tree(df)
            # 在布局完成后绘制模型
            QTimer.singleShot(100, self.draw_segments)  # 延迟100毫秒确保UI完全初始化

    def parse_df_and_populate_tree(self, df):
        """Parse dataframe and populate the tree widget with hierarchical structure."""
        self.tree.clear()  # 清空树
        self.segment_shapes = []
        self.segments = []
        self.ais_shapes = {}  # 清空AIS形状字典
        
        # 存储唯一节点信息
        self.unique_nodes = {}  # 使用ref作为键，(x, y, z)作为值
        self.node_shapes = []  # 存储节点的形状对象
        
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
            # 更新进度条
            progress.setValue(index)
            QCoreApplication.processEvents()
            
            # 如果用户取消了操作
            if progress.wasCanceled():
                break
            
            link_name = row['Link Name']
            ref_origine = row['refOrigine']
            x_origine = float(row['Xorigine'])
            y_origine = float(row['Yorigine'])
            z_origine = float(row['Zorigine'])
            ref_extremite = row['RefExtremite']
            x_extremite = float(row['Xextremite'])
            y_extremite = float(row['Yextremite'])
            z_extremite = float(row['Zextremite'])
            length = row['Length']
            density = row['Density']
            safety = row['Safety']
            route = row['Route']
            action_number = row['Action Number']
            section = row['Section']

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
            
            # 存储当前线段的索引
            segment_index = len(self.segments) - 1
            link_item.setData(0, Qt.UserRole, segment_index)
            
            # 将索引添加到对应的section索引列表中
            section_indices[section].append(segment_index)
    
        # 现在添加节点到树中
        node_indices = []
        for i, (node_id, node_pos) in enumerate(self.unique_nodes.items()):
            node_item = QTreeWidgetItem(nodes_root)
            node_item.setText(0, f"Node: {node_id}")
            
            # 使用特殊格式存储节点ID，区分节点和线段
            node_item.setData(0, Qt.UserRole, f"node_{i}")
            node_indices.append(f"node_{i}")
            
            # 添加节点坐标信息
            x_item = QTreeWidgetItem(node_item)
            x_item.setText(0, f"X= {node_pos[0]}")
            y_item = QTreeWidgetItem(node_item)
            y_item.setText(0, f"Y= {node_pos[1]}")
            z_item = QTreeWidgetItem(node_item)
            z_item.setText(0, f"Z= {node_pos[2]}")
        
        # 设置节点根节点的索引列表
        nodes_root.setData(0, Qt.UserRole, node_indices)
        
        # 设置每个section组的索引列表
        for section, section_item in section_groups.items():
            section_item.setData(0, Qt.UserRole, section_indices[section])
        
        # 将所有索引添加到根节点 (包括线段和节点)
        all_indices = []
        for indices in section_indices.values():
            all_indices.extend(indices)
        all_indices.extend(node_indices)
        main_root.setData(0, Qt.UserRole, all_indices)
            
        # 关闭进度对话框
        progress.setValue(len(df))
        
        # 显示成功消息
        self.show_success_message(f"Excel数据解析完成! 提取了 {len(self.unique_nodes)} 个唯一节点。")
        
        # 展开第一层节点
        self.tree.expandItem(main_root)

    def create_node_shapes(self):
        """创建代表节点的球体形状"""
        self.node_shapes = []  # 重置节点形状列表
        
        # 设置球体的半径
        radius = 40.0  # 球体的半径，可以根据需要调整
        
        for i, (node_id, node_pos) in enumerate(self.unique_nodes.items()):
            # 创建球体
            center = gp_Pnt(*node_pos)
            sphere = BRepPrimAPI_MakeSphere(center, radius).Shape()
            self.node_shapes.append(sphere)
            
            # 创建AIS对象
            ais_shape = AIS_Shape(sphere)
            
            # 设置节点颜色为红色
            color = Quantity_Color(Quantity_NOC_RED)
            self.viewer._display.Context.SetColor(ais_shape, color, False)
            
            # 存储AIS对象，使用 'node_' 前缀区分节点
            self.ais_shapes[f'node_{i}'] = ais_shape

    def draw_segments(self):
        """Draw all segments in the 3D viewer with improved handling of AIS shapes."""
        self.viewer._display.EraseAll()  # 清除现有显示
        self.segment_shapes = []  # 重置形状列表
        self.ais_shapes = {}  # 清空AIS形状字典
        self.highlighted_shapes = []  # 清空高亮列表
        
        # 首先创建节点形状
        self.create_node_shapes()
        
        # 只在首次绘制时显示进度对话框
        progress = None
        total_shapes = len(self.segments) + len(self.node_shapes)
        if self.first_draw and total_shapes > 0:
            progress = QProgressDialog("正在绘制线段和节点...", "取消", 0, total_shapes, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
            progress.show()
            QCoreApplication.processEvents()
        
        # 绘制节点
        for i, sphere in enumerate(self.node_shapes):
            # 更新进度条
            if progress:
                progress.setValue(i)
                QCoreApplication.processEvents()
                
                # 如果用户取消了操作
                if progress.wasCanceled():
                    break
            
            # 获取AIS对象
            ais_shape = self.ais_shapes[f'node_{i}']
            
            # 显示形状
            self.viewer._display.Context.Display(ais_shape, False)
        
        # 绘制线段
        for i, (start, end) in enumerate(self.segments):
            # 更新进度条
            if progress:
                progress.setValue(len(self.node_shapes) + i)
                QCoreApplication.processEvents()
                
                # 如果用户取消了操作
                if progress.wasCanceled():
                    break
                
            start_pnt = gp_Pnt(*start)
            end_pnt = gp_Pnt(*end)
            # 计算圆柱体的参数
            height = start_pnt.Distance(end_pnt)  # 圆柱体的高度（线段的长度）
            radius = 30.0  # 圆柱体的半径（线段的粗细）

            # 计算圆柱体的方向
            direction = gp_Dir(end_pnt.X() - start_pnt.X(),
                               end_pnt.Y() - start_pnt.Y(),
                               end_pnt.Z() - start_pnt.Z())
            axis = gp_Ax2(start_pnt, direction)

            # 创建圆柱体
            cylinder = BRepPrimAPI_MakeCylinder(axis, radius, height).Shape()
            self.segment_shapes.append(cylinder)
            
            # 使用AIS_Shape创建可视化对象
            ais_shape = AIS_Shape(cylinder)
            
            # 设置默认颜色为蓝色
            color = Quantity_Color(Quantity_NOC_BLUE)
            self.viewer._display.Context.SetColor(ais_shape, color, False)
            
            # 显示形状
            self.viewer._display.Context.Display(ais_shape, False)
            
            # 存储AIS对象
            self.ais_shapes[i] = ais_shape

        # 关闭进度对话框并适配视图
        if progress:
            progress.setValue(total_shapes)
            
        self.viewer._display.FitAll()
        self.viewer._display.Repaint()  # 确保重绘视图
        
        # 设置为非首次绘制
        self.first_draw = False

    def draw_imported_shapes(self, show_progress=True):
        """显示导入的STEP形状，使用AIS_Shape进行更好的颜色管理。"""
        self.viewer._display.EraseAll()  # 清除现有显示
        self.ais_shapes = {}  # 清空AIS形状字典
        self.highlighted_shapes = []  # 清空高亮列表
        
        # 只在指定时显示进度对话框
        progress = None
        if show_progress and self.step_shapes:
            total_shapes = len(self.step_shapes)
            progress = QProgressDialog("正在显示导入的形状...", "取消", 0, total_shapes, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
            progress.show()
            QCoreApplication.processEvents()
        
        # 显示所有导入的形状
        for i, (shape_id, shape) in enumerate(self.step_shapes.items()):
            # 更新进度条
            if progress:
                progress.setValue(i)
                QCoreApplication.processEvents()
                
                # 如果用户取消了操作
                if progress.wasCanceled():
                    break
            
            try:
                # 使用AIS_Shape创建可视化对象
                ais_shape = AIS_Shape(shape)
                
                # 设置默认颜色为蓝色
                color = Quantity_Color(Quantity_NOC_BLUE)
                self.viewer._display.Context.SetColor(ais_shape, color, False)
                
                # 显示形状
                self.viewer._display.Context.Display(ais_shape, False)
                
                # 存储AIS对象
                self.ais_shapes[shape_id] = ais_shape
            except Exception as e:
                print(f"无法显示形状 {shape_id}: {e}")
                continue
        
        # 关闭进度对话框并适配视图
        if progress:
            progress.setValue(len(self.step_shapes))
            
        self.viewer._display.FitAll()
        self.viewer._display.Repaint()  # 确保重绘视图

    def highlight_shapes(self, shape_ids, highlight=True):
        """高亮显示或取消高亮指定的形状。"""
        # 确保shape_ids是列表
        if not isinstance(shape_ids, list):
            shape_ids = [shape_ids]
            
        # 分批处理以提高性能
        batch_size = 100  # 每批处理的形状数量
        for i in range(0, len(shape_ids), batch_size):
            batch = shape_ids[i:i + batch_size]
            
            for shape_id in batch:
                if shape_id in self.ais_shapes:
                    if highlight:
                        # 检查是否是节点
                        if isinstance(shape_id, str) and shape_id.startswith('node_'):
                            # 设置为绿色
                            color = Quantity_Color(Quantity_NOC_GREEN)
                        else:
                            # 设置为黄色
                            color = Quantity_Color(Quantity_NOC_YELLOW)
                    else:
                        # 检查是否是节点
                        if isinstance(shape_id, str) and shape_id.startswith('node_'):
                            # 还原为红色
                            color = Quantity_Color(Quantity_NOC_RED)
                        else:
                            # 还原为蓝色
                            color = Quantity_Color(Quantity_NOC_BLUE)
                    
                    self.viewer._display.Context.SetColor(self.ais_shapes[shape_id], color, False)
            
            # 每批次后更新一次显示
            if i + batch_size >= len(shape_ids) or i == 0:
                self.viewer._display.Context.UpdateCurrentViewer()

    def on_tree_item_clicked(self, item, column):
        """使用优化的高亮方法，只更改颜色而不重新绘制形状。"""
        # 获取项目的数据 - 可能是形状ID或索引
        shape_data = item.data(0, Qt.UserRole)
        
        # 处理形状ID列表或单个ID
        if shape_data is None:
            return
            
        # 如果已经选中，再次点击表示取消
        if self.selected_item == item:
            # 取消高亮
            self.highlight_shapes(self.highlighted_shapes, False)
            self.highlighted_shapes = []
            self.selected_item = None
            return
            
        # 先取消之前的高亮
        if self.highlighted_shapes:
            self.highlight_shapes(self.highlighted_shapes, False)
            self.highlighted_shapes = []
            
        # 设置新的选中项
        self.selected_item = item
        
        # 处理节点高亮 (节点ID格式: 'node_X')
        if isinstance(shape_data, str) and shape_data.startswith('node_') and shape_data in self.ais_shapes:
            # 高亮单个节点
            self.highlight_shapes([shape_data], True)
            self.highlighted_shapes = [shape_data]
        # 处理STEP形状
        elif isinstance(shape_data, str) and shape_data in self.ais_shapes:
            # 高亮单个形状
            self.highlight_shapes([shape_data], True)
            self.highlighted_shapes = [shape_data]
        elif isinstance(shape_data, list):
            # 对于字符串列表(STEP形状ID、节点ID)或整数列表(线段索引)
            valid_ids = []
            for id in shape_data:
                if id in self.ais_shapes:
                    valid_ids.append(id)
            if valid_ids:
                self.highlight_shapes(valid_ids, True)
                self.highlighted_shapes = valid_ids
        elif isinstance(shape_data, int) and shape_data in self.ais_shapes:
            # 高亮单个线段
            self.highlight_shapes([shape_data], True)
            self.highlighted_shapes = [shape_data]

    def closeEvent(self, event):
        """窗口关闭时清理资源"""
        try:
            # 1. 清除所有显示的图形
            if hasattr(self, 'viewer') and self.viewer:
                self.viewer._display.Context.EraseAll(True)  # 立即更新视图

            # 2. 关闭并释放 viewer 对象
            if hasattr(self, 'viewer') and self.viewer:
                self.viewer.close()  # 关闭 viewer
                self.viewer.deleteLater()  # 确保 viewer 对象被释放
                self.viewer = None  # 将 viewer 引用置为 None

            # 3. 清理计时器
            if hasattr(self, 'timer') and self.timer:
                self.timer.stop()
                # 如果计时器已连接，断开连接
                if hasattr(self, 'timer_connected') and self.timer_connected:
                    try:
                        self.timer.timeout.disconnect()
                    except:
                        pass
                    self.timer_connected = False
            
            print("资源清理完成")
        except Exception as e:
            print(f"清理资源时发生错误: {e}")
        finally:
            # 确保调用父类的 closeEvent
            super().closeEvent(event)

    def import_file(self):
        """根据选择的格式导入文件"""
        file_format = self.file_format_combo.currentText()
        
        if file_format == "STEP":
            self.import_step()
        elif file_format == "IGES":
            self.import_iges()
        else:
            QMessageBox.warning(self, "格式错误", f"不支持的文件格式: {file_format}")

    def export_file(self):
        """根据选择的格式导出文件"""
        file_format = self.export_format_combo.currentText()
        
        if file_format == "STEP":
            self.export_to_step()
        elif file_format == "IGES":
            self.export_to_iges()
        else:
            QMessageBox.warning(self, "格式错误", f"不支持的文件格式: {file_format}")

    def export_to_step(self):
        """导出 3D 模型为 STEP 文件"""
        file_path, _ = QFileDialog.getSaveFileName(self, "保存为 STEP 文件", "", "STEP 文件 (*.step *.stp)")
        if not file_path:
            return
            
        # 显示进度对话框
        progress = QProgressDialog("正在导出 STEP 文件...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
        progress.show()
        progress.setValue(10)  # 设置初始进度
        QCoreApplication.processEvents()
            
        try:
            step_writer = STEPControl_Writer()
            Interface_Static_SetCVal("write.step.schema", "AP203")
            
            progress.setValue(30)
            QCoreApplication.processEvents()
            
            # 如果有导入的STEP模型
            if self.main_shape is not None:
                progress.setLabelText("正在导出主模型...")
                step_writer.Transfer(self.main_shape, STEPControl_AsIs)
            # 如果有导入的形状
            elif self.step_shapes:
                progress.setLabelText("正在导出所有形状...")
                # 创建一个组合体，将所有形状添加进去
                from OCC.Core.TopoDS import TopoDS_Compound
                from OCC.Core.BRep import BRep_Builder
                
                compound = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(compound)
                
                total_shapes = len(self.step_shapes)
                for i, (shape_id, shape) in enumerate(self.step_shapes.items()):
                    # 添加到组合体
                    builder.Add(compound, shape)
                    # 更新进度，从30%到80%
                    progress_value = 30 + int(50 * (i + 1) / total_shapes)
                    progress.setValue(progress_value)
                    QCoreApplication.processEvents()
                    
                    # 如果用户取消了操作
                    if progress.wasCanceled():
                        return
                
                # 导出组合体
                step_writer.Transfer(compound, STEPControl_AsIs)
            else:
                # 导出创建的圆柱体和球体
                progress.setLabelText("正在导出线段和节点...")
                # 创建一个组合体，将所有圆柱体和球体添加进去
                from OCC.Core.TopoDS import TopoDS_Compound
                from OCC.Core.BRep import BRep_Builder
                
                compound = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(compound)
                
                # 添加所有线段形状
                total_shapes = len(self.segment_shapes) + len(self.node_shapes)
                shape_counter = 0
                
                for shape in self.segment_shapes:
                    # 添加到组合体
                    builder.Add(compound, shape)
                    # 更新进度，从30%到80%
                    shape_counter += 1
                    progress_value = 30 + int(50 * shape_counter / total_shapes)
                    progress.setValue(progress_value)
                    QCoreApplication.processEvents()
                    
                    # 如果用户取消了操作
                    if progress.wasCanceled():
                        return
                
                # 添加所有节点形状
                for shape in self.node_shapes:
                    # 添加到组合体
                    builder.Add(compound, shape)
                    # 更新进度，从30%到80%
                    shape_counter += 1
                    progress_value = 30 + int(50 * shape_counter / total_shapes)
                    progress.setValue(progress_value)
                    QCoreApplication.processEvents()
                    
                    # 如果用户取消了操作
                    if progress.wasCanceled():
                        return
                
                # 导出组合体
                step_writer.Transfer(compound, STEPControl_AsIs)
            
            progress.setValue(80)
            progress.setLabelText("正在写入文件...")
            QCoreApplication.processEvents()
            
            # 写入文件
            status = step_writer.Write(file_path)
            
            progress.setValue(100)
            
            if status == IFSelect_RetDone:
                self.show_success_message(f"已成功导出到 {file_path}")
            else:
                QMessageBox.warning(self, "导出失败", "导出过程中出现错误，请重试。")
                
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "导出错误", f"导出过程中出现异常: {str(e)}")
            print(f"导出错误: {e}")
            import traceback
            traceback.print_exc()
    
    def export_to_iges(self):
        """导出 3D 模型为 IGES 文件"""
        file_path, _ = QFileDialog.getSaveFileName(self, "保存为 IGES 文件", "", "IGES 文件 (*.iges *.igs)")
        if not file_path:
            return
            
        # 显示进度对话框
        progress = QProgressDialog("正在导出 IGES 文件...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
        progress.show()
        progress.setValue(10)  # 设置初始进度
        QCoreApplication.processEvents()
            
        try:
            iges_writer = IGESControl_Writer()
            
            progress.setValue(30)
            QCoreApplication.processEvents()
            
            # 如果有导入的STEP模型
            if self.main_shape is not None:
                progress.setLabelText("正在导出主模型...")
                iges_writer.AddShape(self.main_shape)
            # 如果有导入的形状
            elif self.step_shapes:
                progress.setLabelText("正在导出所有形状...")
                # 创建一个组合体，将所有形状添加进去
                from OCC.Core.TopoDS import TopoDS_Compound
                from OCC.Core.BRep import BRep_Builder
                
                compound = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(compound)
                
                total_shapes = len(self.step_shapes)
                for i, (shape_id, shape) in enumerate(self.step_shapes.items()):
                    # 添加到组合体
                    builder.Add(compound, shape)
                    # 更新进度，从30%到80%
                    progress_value = 30 + int(50 * (i + 1) / total_shapes)
                    progress.setValue(progress_value)
                    QCoreApplication.processEvents()
                    
                    # 如果用户取消了操作
                    if progress.wasCanceled():
                        return
                
                # 导出组合体
                iges_writer.AddShape(compound)
            else:
                # 导出创建的圆柱体和球体
                progress.setLabelText("正在导出线段和节点...")
                # 创建一个组合体，将所有圆柱体和球体添加进去
                from OCC.Core.TopoDS import TopoDS_Compound
                from OCC.Core.BRep import BRep_Builder
                
                compound = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(compound)
                
                # 添加所有线段形状和节点形状
                total_shapes = len(self.segment_shapes) + len(self.node_shapes)
                shape_counter = 0
                
                for shape in self.segment_shapes:
                    # 添加到组合体
                    builder.Add(compound, shape)
                    # 更新进度，从30%到80%
                    shape_counter += 1
                    progress_value = 30 + int(50 * shape_counter / total_shapes)
                    progress.setValue(progress_value)
                    QCoreApplication.processEvents()
                    
                    # 如果用户取消了操作
                    if progress.wasCanceled():
                        return
                
                # 添加所有节点形状
                for shape in self.node_shapes:
                    # 添加到组合体
                    builder.Add(compound, shape)
                    # 更新进度，从30%到80%
                    shape_counter += 1
                    progress_value = 30 + int(50 * shape_counter / total_shapes)
                    progress.setValue(progress_value)
                    QCoreApplication.processEvents()
                    
                    # 如果用户取消了操作
                    if progress.wasCanceled():
                        return
                
                # 导出组合体
                iges_writer.AddShape(compound)
            
            progress.setValue(80)
            progress.setLabelText("正在写入文件...")
            QCoreApplication.processEvents()
            
            # 写入文件
            status = iges_writer.Write(file_path)
            
            progress.setValue(100)
            
            if status == IFSelect_RetDone:
                self.show_success_message(f"已成功导出到 {file_path}")
            else:
                QMessageBox.warning(self, "导出失败", "导出过程中出现错误，请重试。")
                
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "导出错误", f"导出过程中出现异常: {str(e)}")
            print(f"导出错误: {e}")
            import traceback
            traceback.print_exc()

    def import_step(self):
        """导入STEP文件并解析"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 STEP 文件", "", "STEP 文件 (*.step *.stp)")
        if not file_path:
            return
            
        self.import_cad_file(file_path, "STEP")
    
    def import_iges(self):
        """导入IGES文件并解析"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 IGES 文件", "", "IGES 文件 (*.iges *.igs)")
        if not file_path:
            return
            
        self.import_cad_file(file_path, "IGES")

    def import_cad_file(self, file_path, file_format):
        """通用CAD文件导入函数"""
        # 显示进度对话框
        progress = QProgressDialog(f"正在导入 {file_format} 文件...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
        progress.show()
        progress.setValue(10)  # 设置初始进度
        QCoreApplication.processEvents()
            
        try:
            # 清空现有数据
            self.tree.clear()
            self.segment_shapes = []
            self.segments = []
            self.step_shapes = {}
            self.ais_shapes = {}
            self.highlighted_shapes = []
            self.unique_nodes = {}  # 清空节点数据
            self.node_shapes = []   # 清空节点形状
            self.viewer._display.EraseAll()
            
            progress.setValue(20)
            QCoreApplication.processEvents()
            
            # 根据格式选择合适的读取器
            if file_format == "STEP":
                reader = STEPControl_Reader()
            elif file_format == "IGES":
                reader = IGESControl_Reader()
            else:
                raise ValueError(f"不支持的文件格式: {file_format}")
            
            status = reader.ReadFile(file_path)
            
            progress.setValue(40)
            QCoreApplication.processEvents()
            
            if status != IFSelect_RetDone:
                progress.close()
                QMessageBox.warning(self, "导入错误", f"无法读取{file_format}文件，格式可能不支持。")
                return
            
            # 转换根形状
            reader.TransferRoots()
            
            # 获取形状
            self.main_shape = reader.Shape()
            
            if self.main_shape.IsNull():
                progress.close()
                QMessageBox.warning(self, "导入错误", f"{file_format}文件中没有找到有效的形状。")
                return
                
            progress.setValue(60)
            QCoreApplication.processEvents()
            
            # 创建树的根节点
            file_basename = file_path.split("/")[-1]
            root = QTreeWidgetItem(self.tree)
            root.setText(0, f"{file_format} Model: {file_basename}")
            
            # 解析形状并填充树
            self.analyze_shape_and_build_tree(self.main_shape, root, progress)
            
            # 展开根节点
            self.tree.expandItem(root)
            
            progress.setValue(90)
            
            # 重置首次绘制标志，以便在导入后的第一次绘制时显示进度条
            self.first_draw = True
            
            # 显示导入的形状
            self.draw_imported_shapes(show_progress=True)
            
            progress.setValue(100)
            
            # 计算总元素数
            total_elements = len(self.step_shapes)
            self.show_success_message(f"已成功导入{file_format}文件，包含 {total_elements} 个元素。")
            
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "导入错误", f"导入过程中出现异常: {str(e)}")
            print(f"导入错误详情: {e}", flush=True)
            import traceback
            traceback.print_exc()

    def analyze_shape_and_build_tree(self, shape, parent_item, progress=None):
        """分析形状的层次结构并构建树视图 - 优化版本，修复了Copy()和索引错误"""
        if progress:
            progress.setValue(70)
            QCoreApplication.processEvents()
            
        # 记录要保存的所有形状ID列表
        all_shape_ids = []
        
        # 快速创建分类节点，减少树操作
        shape_types = {
            TopAbs_SOLID: "实体",
            TopAbs_SHELL: "壳体", 
            TopAbs_FACE: "面",
            TopAbs_WIRE: "线框",
            TopAbs_EDGE: "边缘"
        }
        
        # 预处理 - 先收集所有形状，减少树的操作次数
        shape_collections = {t: [] for t in shape_types}
        
        # 收集各类型的形状
        for shape_type in shape_types:
            explorer = TopExp_Explorer(shape, shape_type)
            while explorer.More():
                # 直接使用当前形状，不尝试复制
                current_shape = explorer.Current()
                shape_collections[shape_type].append(current_shape)
                explorer.Next()
                
        # 创建树节点
        for shape_type, type_name in shape_types.items():
            shapes = shape_collections[shape_type]
            if shapes:
                # 创建类型节点
                type_node = QTreeWidgetItem(parent_item)
                type_node.setText(0, f"{type_name} ({len(shapes)})")
                
                # 记录这个类型下的所有形状ID
                type_shape_ids = []
                
                # 添加所有形状 - 修正批次处理逻辑
                batch_size = 50
                num_batches = (len(shapes) + batch_size - 1) // batch_size  # 向上取整
                
                for batch_index in range(num_batches):
                    # 计算这个批次的开始和结束索引
                    start_idx = batch_index * batch_size
                    end_idx = min(start_idx + batch_size, len(shapes))
                    
                    # 创建批次节点
                    batch_node = QTreeWidgetItem(type_node)
                    batch_node.setText(0, f"{type_name} {start_idx + 1} - {end_idx}")
                    
                    # 处理这个批次中的所有形状
                    batch_shape_ids = []
                    for i in range(start_idx, end_idx):
                        # 生成唯一ID
                        current_shape = shapes[i]
                        shape_id = f"{type_name.lower()}_{i+1}_{id(current_shape) % 10000}"  # 简化ID
                        
                        # 保存形状和ID
                        self.step_shapes[shape_id] = current_shape
                        batch_shape_ids.append(shape_id)
                        type_shape_ids.append(shape_id)
                        all_shape_ids.append(shape_id)
                    
                    # 将当前批次的形状ID存储到批次节点
                    batch_node.setData(0, Qt.UserRole, batch_shape_ids)
                
                # 将所有形状ID存储到类型节点
                type_node.setData(0, Qt.UserRole, type_shape_ids)
        
        # 将所有形状ID存储到父节点
        parent_item.setData(0, Qt.UserRole, all_shape_ids)
        
        # 保存主形状用于导出
        self.main_shape = shape

    def get_shape_type_name(self, shape):
        """获取形状类型的名称"""
        if shape.ShapeType() == TopAbs_EDGE:
            return "边缘"
        elif shape.ShapeType() == TopAbs_VERTEX:
            return "顶点"
        elif shape.ShapeType() == TopAbs_FACE:
            return "面"
        elif shape.ShapeType() == TopAbs_SOLID:
            return "实体"
        elif shape.ShapeType() == TopAbs_SHELL:
            return "壳"
        elif shape.ShapeType() == TopAbs_COMPOUND:
            return "组合体"
        elif shape.ShapeType() == TopAbs_COMPSOLID:
            return "复合实体"
        elif shape.ShapeType() == TopAbs_WIRE:
            return "线框"
        else:
            return f"形状 (类型 {shape.ShapeType()})"
            
    def show_success_message(self, message):
        """显示成功消息，并在1秒后自动关闭"""
        # 如果已有消息框正在显示，先关闭它
        if hasattr(self, 'msg_box') and self.msg_box:
            self.msg_box.close()
            self.msg_box = None
            
        # 创建并显示新的消息框
        self.msg_box = QMessageBox(self)
        self.msg_box.setWindowTitle("成功")
        self.msg_box.setText(message)
        self.msg_box.setIcon(QMessageBox.Information)
        self.msg_box.setStandardButtons(QMessageBox.Ok)
        self.msg_box.show()
        
        # 设置定时器在1秒后自动关闭消息框
        if hasattr(self, 'timer'):
            # 如果已经连接了信号，先断开
            if hasattr(self, 'timer_connected') and self.timer_connected:
                try:
                    self.timer.timeout.disconnect()
                except TypeError:
                    # 忽略断开失败的错误
                    pass
                
            # 重新连接信号
            self.timer.timeout.connect(self.close_message_box)
            self.timer_connected = True
            self.timer.start(1000)  # 1000毫秒 = 1秒
        
    def close_message_box(self):
        """关闭当前显示的消息框"""
        if hasattr(self, 'msg_box') and self.msg_box:
            self.msg_box.close()
            self.msg_box = None
        
        # 断开信号连接
        if hasattr(self, 'timer_connected') and self.timer_connected:
            try:
                self.timer.timeout.disconnect()
            except TypeError:
                # 忽略断开失败的错误
                pass
            self.timer_connected = False


def main(xlsx_file=None):
    app = QApplication(sys.argv)
    
    # 如果提供了Excel文件参数，则加载它
    if xlsx_file:
        df = pd.read_excel(xlsx_file)
        window = MainWindow(df)
    else:
        window = MainWindow()

    # 设置窗口大小
    window.resize(1000, 800)  # 调整窗口大小为 1000x800

    # 设置窗口居中
    screen_geometry = QDesktopWidget().screenGeometry()  # 获取屏幕的几何信息
    x = (screen_geometry.width() - window.width()) // 2  # 计算窗口居中的 x 坐标
    y = (screen_geometry.height() - window.height()) // 2  # 计算窗口居中的 y 坐标
    window.move(x, y)  # 移动窗口到居中位置

    window.show()
    app.exec_()


if __name__ == "__main__":
    # 使用 argparse 解析命令行参数
    parser = argparse.ArgumentParser(description="航电布线可视化系统")
    parser.add_argument("xlsx_file", type=str, nargs='?', default=None, help="EXCEL 文件路径")
    args = parser.parse_args()

    # 调用主函数
    main(args.xlsx_file)