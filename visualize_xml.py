import sys
import argparse
import xml.etree.ElementTree as ET
from OCC.Core.Quantity import Quantity_Color
from OCC.Core._Quantity import Quantity_TOC_RGB
from OCC.Display.backend import load_backend
from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_Reader, STEPControl_AsIs
from OCC.Core.IGESControl import IGESControl_Reader, IGESControl_Writer
from OCC.Core.Interface import Interface_Static_SetCVal
from OCC.Core.IFSelect import IFSelect_RetDone

load_backend("qt-pyqt5")
from OCC.Display.qtDisplay import qtViewer3d

from PyQt5.QtWidgets import (
    QApplication, QTreeWidget, QTreeWidgetItem, QWidget, QMainWindow,
    QHBoxLayout, QVBoxLayout, QDesktopWidget, QPushButton, QFileDialog,
    QLabel, QComboBox, QGroupBox, QMessageBox, QProgressDialog
)
from PyQt5.QtCore import Qt, QTimer, QCoreApplication
from PyQt5.QtGui import QFont
from OCC.Core.BRepPrimAPI import BRepPrimAPI_MakeCylinder
from OCC.Core.gp import gp_Pnt, gp_Dir, gp_Ax2
from OCC.Core.TopoDS import TopoDS_Shape, TopoDS_Compound
from OCC.Core.BRep import BRep_Builder
from OCC.Core.AIS import AIS_Shape
from OCC.Core.Quantity import Quantity_NOC_BLUE, Quantity_NOC_YELLOW


class MainWindow(QWidget):
    def __init__(self, xml_file=None):
        super().__init__()
        self.setWindowTitle("航电布线可视化系统")
        
        # 初始化消息框和计时器
        self.msg_box = None
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer_connected = False
        
        # 高亮和颜色管理
        self.ais_shapes = {}  # 存储所有AIS对象，用于颜色管理
        self.highlighted_shapes = []  # 当前高亮的形状IDs
        
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
        self.tree.setHeaderLabels(["模型结构"])  # 移除详细信息列
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
        
        # 保存导入的STEP/IGES模型
        self.step_shapes = {}  # 形状ID到形状对象的映射
        self.main_shape = None
        
        # 树项点击事件
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        self.selected_item = None

        # 如果有XML文件，则解析并显示
        if xml_file is not None:
            self.parse_xml_and_populate_tree(xml_file)
            # 在布局完成后绘制模型
            QTimer.singleShot(100, self.draw_segments)  # 延迟100毫秒确保UI完全初始化

    def parse_xml_and_populate_tree(self, file_path):
        """解析XML并填充树状结构"""
        self.tree.clear()  # 清空树
        self.segment_shapes = []
        self.segments = []
        self.ais_shapes = {}  # 清空AIS形状字典
        
        # 显示进度对话框
        progress = QProgressDialog("正在解析XML数据...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
        progress.show()
        QCoreApplication.processEvents()
        
        try:
            # 解析XML树
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            progress.setValue(20)
            QCoreApplication.processEvents()
            
            # 计算大概的总网络数量，用于进度显示
            total_networks = sum(1 for _ in root.findall('.//Network'))
            current_network = 0
            
            for net in root.findall('Net'):
                net_item = QTreeWidgetItem(self.tree)
                net_item.setText(0, net.get('name'))
                net_network_lst = []
                
                # 处理设备
                devices = net.find('Devices')
                if devices is not None:
                    devices_item = QTreeWidgetItem(net_item)
                    devices_item.setText(0, "Equloc")
                    
                    # 更新进度
                    progress.setValue(30)
                    progress.setLabelText("正在处理设备...")
                    QCoreApplication.processEvents()
                    
                    for device in devices.findall('Device'):
                        device_item = QTreeWidgetItem(devices_item)
                        device_item.setText(0, device.get('name'))
                        # 不再添加详细信息列
                
                # 处理等电位点
                isoelectricPoints = net.find('IsoelectricPoints')
                if isoelectricPoints is not None:
                    isoelectricPoints_item = QTreeWidgetItem(net_item)
                    isoelectricPoints_item.setText(0, "IsoelectricPoints")
                    
                    # 更新进度
                    progress.setValue(40)
                    progress.setLabelText("正在处理等电位点...")
                    QCoreApplication.processEvents()
                    
                    for isoelePt in isoelectricPoints.findall('IsoelePt'):
                        isoelePt_item = QTreeWidgetItem(isoelectricPoints_item)
                        isoelePt_item.setText(0, isoelePt.get('name'))
                
                # 更新进度
                progress.setValue(50)
                progress.setLabelText("正在处理子网络...")
                QCoreApplication.processEvents()
                
                # 处理子网络
                subnet_count = len(net.findall('SubNet'))
                for subnet_idx, subnet in enumerate(net.findall('SubNet')):
                    subnet_item = QTreeWidgetItem(net_item)
                    subnet_item.setText(0, subnet.get('name'))
                    subnet_network_lst = []
                    
                    # 更新子网络进度
                    subnet_progress = 50 + int(30 * (subnet_idx + 1) / subnet_count)
                    progress.setValue(subnet_progress)
                    progress.setLabelText(f"正在处理子网络 {subnet_idx + 1}/{subnet_count}...")
                    QCoreApplication.processEvents()
                    
                    # 处理网络起点和终点
                    netStartPoint = subnet.find('NetStartPoint')
                    netEndPoint = subnet.find('NetEndPoint')
                    
                    if netStartPoint is not None and netEndPoint is not None:
                        net_start = (
                            float(netStartPoint.get('X')),
                            float(netStartPoint.get('Y')),
                            float(netStartPoint.get('Z'))
                        )
                        net_end = (
                            float(netEndPoint.get('X')),
                            float(netEndPoint.get('Y')),
                            float(netEndPoint.get('Z'))
                        )
                        
                        netStartPoint_item = QTreeWidgetItem(subnet_item)
                        netStartPoint_item.setText(0, "NetStartPoint")
                        netEndPoint_item = QTreeWidgetItem(subnet_item)
                        netEndPoint_item.setText(0, "NetEndPoint")
                        
                        net_start_point_item = QTreeWidgetItem(netStartPoint_item)
                        net_start_point_item.setText(0, netStartPoint.get('name'))
                        
                        net_end_point_item = QTreeWidgetItem(netEndPoint_item)
                        net_end_point_item.setText(0, netEndPoint.get('name'))
                    
                    # 处理段
                    for segment in subnet.findall('Segement'):
                        segment_item = QTreeWidgetItem(subnet_item)
                        segment_item.setText(0, segment.get('name'))
                        network_lst = []
                        
                        # 处理网络
                        for network in segment.findall('Network'):
                            current_network += 1
                            network_item = QTreeWidgetItem(segment_item)
                            network_item.setText(0, network.get('name'))
                            
                            start_point = network.find('StartPoint')
                            end_point = network.find('EndPoint')
                            
                            if start_point is not None and end_point is not None:
                                start = (
                                    float(start_point.get('x')),
                                    float(start_point.get('y')),
                                    float(start_point.get('z'))
                                )
                                end = (
                                    float(end_point.get('x')),
                                    float(end_point.get('y')),
                                    float(end_point.get('z'))
                                )
                                self.segments.append((start, end))
                                
                                start_point_item = QTreeWidgetItem(network_item)
                                start_point_item.setText(0, start_point.get('name'))
                                
                                end_point_item = QTreeWidgetItem(network_item)
                                end_point_item.setText(0, end_point.get('name'))
                                
                                network_item.setData(0, Qt.UserRole, len(self.segments) - 1)  # 存储索引
                                network_lst.append(len(self.segments) - 1)
                            
                            # 更新网络进度
                            if total_networks > 0:
                                network_progress = 50 + int(40 * current_network / total_networks)
                                if network_progress % 5 == 0:  # 不要更新太频繁
                                    progress.setValue(network_progress)
                                    QCoreApplication.processEvents()
                        
                        segment_item.setData(0, Qt.UserRole, network_lst)
                        subnet_network_lst.extend(network_lst)
                    
                    subnet_item.setData(0, Qt.UserRole, subnet_network_lst)
                    net_network_lst.extend(subnet_network_lst)
                
                net_item.setData(0, Qt.UserRole, net_network_lst)
            
            progress.setValue(100)
            QCoreApplication.processEvents()
            
            # 展开顶级节点
            self.tree.expandToDepth(0)
            
            # 显示成功消息
            self.show_success_message("XML数据解析完成!")

        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "解析错误", f"解析XML时出错: {str(e)}")
            print(f"解析错误: {e}")
            import traceback
            traceback.print_exc()

    def draw_segments(self):
        """使用优化的方式绘制所有线段"""
        self.viewer._display.EraseAll()  # 清除现有显示
        self.segment_shapes = []  # 重置形状列表
        self.ais_shapes = {}  # 清空AIS形状字典
        self.highlighted_shapes = []  # 清空高亮列表
        
        # 显示进度对话框
        if len(self.segments) > 0:
            progress = QProgressDialog("正在绘制线段...", "取消", 0, len(self.segments), self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(500)  # 设置最小显示时间为500ms
            progress.show()
        else:
            progress = None
        
        # 批量创建和显示形状以提高性能
        ais_shapes_batch = []
        batch_size = 50  # 每批处理的线段数量
        
        for i, (start, end) in enumerate(self.segments):
            # 更新进度条
            if progress:
                progress.setValue(i)
                QCoreApplication.processEvents()
                
                # 如果用户取消了操作
                if progress.wasCanceled():
                    break
                
            start_pnt = gp_Pnt(*start)
            end_pnt = gp_Pnt(*end)
            # 计算圆柱体的参数
            height = start_pnt.Distance(end_pnt)  # 圆柱体的高度（线段的长度）
            radius = 15.0  # 圆柱体的半径（线段的粗细）

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
            ais_shapes_batch.append((i, ais_shape))
            
            # 如果达到批处理大小或是最后一个元素，处理当前批次
            if len(ais_shapes_batch) >= batch_size or i == len(self.segments) - 1:
                # 批量设置颜色和显示
                for idx, shape in ais_shapes_batch:
                    # 设置默认颜色为蓝色
                    color = Quantity_Color(Quantity_NOC_BLUE)
                    self.viewer._display.Context.SetColor(shape, color, False)
                    
                    # 显示形状，但不立即更新视图
                    self.viewer._display.Context.Display(shape, False)
                    
                    # 存储AIS对象
                    self.ais_shapes[idx] = shape
                
                # 清空批处理列表
                ais_shapes_batch = []
                # 更新视图
                self.viewer._display.Context.UpdateCurrentViewer()

        # 关闭进度对话框并适配视图
        if progress:
            progress.setValue(len(self.segments))
            
        self.viewer._display.FitAll()
        self.viewer._display.Repaint()  # 确保重绘视图

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
                        # 设置为黄色
                        color = Quantity_Color(Quantity_NOC_YELLOW)
                    else:
                        # 还原为蓝色
                        color = Quantity_Color(Quantity_NOC_BLUE)
                    
                    self.viewer._display.Context.SetColor(self.ais_shapes[shape_id], color, False)
            
            # 每批次后更新一次显示
            if i + batch_size >= len(shape_ids) or i == 0:
                self.viewer._display.Context.UpdateCurrentViewer()

    def on_tree_item_clicked(self, item, column):
        """优化的高亮方法，只更改颜色而不重新绘制形状"""
        # 获取项目的数据 - 可能是形状ID或索引
        shape_data = item.data(0, Qt.UserRole)
        
        # 如果没有数据，直接返回
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
        
        # 处理STEP/IGES形状
        if isinstance(shape_data, str) and shape_data in self.ais_shapes:
            # 高亮单个形状
            self.highlight_shapes(shape_data, True)
            self.highlighted_shapes = [shape_data]
        elif isinstance(shape_data, list):
            # 对于线段索引列表
            valid_ids = []
            for id in shape_data:
                if isinstance(id, int) and id in self.ais_shapes:
                    valid_ids.append(id)
            if valid_ids:
                self.highlight_shapes(valid_ids, True)
                self.highlighted_shapes = valid_ids
        elif isinstance(shape_data, int) and shape_data in self.ais_shapes:
            # 高亮单个线段
            self.highlight_shapes(shape_data, True)
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
            
            # 如果有导入的STEP模型，则导出该模型
            if self.main_shape is not None:
                progress.setLabelText("正在导出主模型...")
                step_writer.Transfer(self.main_shape, STEPControl_AsIs)
            # 如果有导入的形状
            elif self.step_shapes:
                progress.setLabelText("正在导出所有形状...")
                # 创建一个组合体，将所有形状添加进去
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
                # 导出创建的圆柱体
                progress.setLabelText("正在导出线段...")
                # 创建一个组合体，将所有圆柱体添加进去
                compound = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(compound)
                
                total_shapes = len(self.segment_shapes)
                for i, shape in enumerate(self.segment_shapes):
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
            
            # 如果有导入的STEP/IGES模型
            if self.main_shape is not None:
                progress.setLabelText("正在导出主模型...")
                iges_writer.AddShape(self.main_shape)
            # 如果有导入的形状
            elif self.step_shapes:
                progress.setLabelText("正在导出所有形状...")
                # 创建一个组合体，将所有形状添加进去
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
                # 导出创建的圆柱体
                progress.setLabelText("正在导出线段...")
                # 创建一个组合体，将所有圆柱体添加进去
                compound = TopoDS_Compound()
                builder = BRep_Builder()
                builder.MakeCompound(compound)
                
                total_shapes = len(self.segment_shapes)
                for i, shape in enumerate(self.segment_shapes):
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
            
            # 预处理设置 - 提高大型文件的导入性能
            if file_format == "STEP":
                Interface_Static_SetCVal("read.step.product.mode", "0")  # 简化模式
                Interface_Static_SetCVal("read.step.product.context", "0")
                Interface_Static_SetCVal("read.step.shape.repr", "1")
            elif file_format == "IGES":
                Interface_Static_SetCVal("read.iges.bspline.continuity", "0")  # 降低曲面连续性要求
                
            progress.setLabelText(f"正在读取{file_format}文件...")
            status = reader.ReadFile(file_path)
            
            progress.setValue(40)
            QCoreApplication.processEvents()
            
            if status != IFSelect_RetDone:
                progress.close()
                QMessageBox.warning(self, "导入错误", f"无法读取{file_format}文件，格式可能不支持。")
                return
            
            # 转换根形状
            progress.setLabelText(f"正在转换{file_format}数据...")
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
            progress.setLabelText(f"正在分析{file_format}模型结构...")
            self.analyze_shape_and_build_tree(self.main_shape, root, progress)
            
            # 展开根节点
            self.tree.expandItem(root)
            
            progress.setValue(90)
            
            # 显示导入的形状
            progress.setLabelText(f"正在显示{file_format}模型...")
            QCoreApplication.processEvents()
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
        """分析形状的层次结构并构建树视图"""
        from OCC.Core.TopAbs import TopAbs_EDGE, TopAbs_FACE, TopAbs_SOLID, TopAbs_SHELL, TopAbs_WIRE
        from OCC.Core.TopExp import TopExp_Explorer
        
        if progress:
            progress.setLabelText("正在分析模型结构...")
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
        
        if progress:
            progress.setValue(80)
            QCoreApplication.processEvents()

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


def main(xml_file=None):
    app = QApplication(sys.argv)
    
    window = MainWindow(xml_file)

    # 设置窗口大小
    window.resize(1600, 800)  # 调整窗口大小为 1600x800

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
    parser.add_argument("xml_file", type=str, nargs='?', default=None, help="XML 文件路径")
    args = parser.parse_args()

    # 调用主函数
    main(args.xml_file)