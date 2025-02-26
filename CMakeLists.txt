cmake_minimum_required(VERSION 3.10)
project(AvionicsWiringVisualization)

set(Qt5_DIR "/opt/homebrew/Cellar/qt@5/5.15.16/lib/cmake/Qt5")
# 查找 QT
find_package(Qt5 COMPONENTS Widgets Xml REQUIRED)

# 设置 C++ 标准为 C++11 或更高
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_AUTOMOC ON)

# 添加可执行文件，包含新的源文件
add_executable(wire_visualization main.cpp)

# 链接 QT 库
target_link_libraries(wire_visualization PRIVATE Qt5::Widgets Qt5::Xml)