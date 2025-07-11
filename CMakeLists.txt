cmake_minimum_required(VERSION 3.15)
project(AviationWiringVisualization LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

set(Qt5_DIR "/opt/homebrew/Cellar/qt@5/5.15.16/lib/cmake/Qt5")
set(OpenCASCADE_DIR "/opt/homebrew/Cellar/opencascade/7.8.1_1/lib/cmake/opencascade")
# 查找所需的包
find_package(Qt5 COMPONENTS Core Widgets OpenGL REQUIRED)
find_package(OpenCASCADE REQUIRED)

# 包含目录
include_directories(${OpenCASCADE_INCLUDE_DIR})
include_directories(${CMAKE_CURRENT_SOURCE_DIR})

# 定义源文件
set(SOURCES
    configmanager.cpp
    filevisualizer.cpp
    main.cpp
)

# 定义头文件
set(HEADERS
    configmanager.h
    filevisualizer.h
)

# 创建可执行文件
add_executable(${PROJECT_NAME} ${SOURCES} ${HEADERS})

# 链接库
target_link_libraries(${PROJECT_NAME} PRIVATE
    Qt5::Core
    Qt5::Widgets
    Qt5::OpenGL
    ${OpenCASCADE_LIBRARIES}
)

# 复制依赖的DLL到输出目录（仅Windows）
if(WIN32)
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
        $<TARGET_FILE:Qt5::Core>
        $<TARGET_FILE:Qt5::Widgets>
        $<TARGET_FILE:Qt5::OpenGL>
        $<TARGET_FILE_DIR:${PROJECT_NAME}>
    )
endif()