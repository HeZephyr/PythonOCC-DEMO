#-------------------------------------------------
# Qt Project File for AviationWiringVisualization
# Based on Qt 5.15+ and OpenCASCADE 7.8.1
#-------------------------------------------------

QT       += core gui widgets opengl
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = AviationWiringVisualization
TEMPLATE = app

# Compiler configurations
CONFIG += c++14
QT_MINOR_VERSION = 15.16
requires(qtVersion() >= 5.15.16)

# OpenCASCADE configuration (macOS Homebrew路径示例)
macx {
    INCLUDEPATH += /opt/homebrew/Cellar/opencascade/7.8.1_1/include/opencascade
    LIBS += -L/opt/homebrew/Cellar/opencascade/7.8.1_1/lib \
        -lTKernel -lTKG2d -lTKG3d -lTKMath \
        -lTKIGES -lTKSTL -lTKSTEP -lTKBRep
}

# 源文件和头文件
SOURCES += \
    configmanager.cpp \
    filevisualizer.cpp \
    main.cpp

HEADERS += \
    configmanager.h \
    filevisualizer.h

# 自动处理Qt元对象系统
FORMS += \
    # 如果有.ui文件在此添加

RESOURCES += \
    # 如果有.qrc文件在此添加

# Windows特定配置
win32 {
    # OpenCASCADE Windows路径示例（需根据实际安装路径修改）
    INCLUDEPATH += C:/OpenCASCADE7.8.1/include
    LIBS += -LC:/OpenCASCADE7.8.1/win64/vc14/lib \
        -lTKernel -lTKG2d -lTKG3d -lTKMath
    
    # 自动复制Qt DLL到构建目录
    QMAKE_POST_LINK += $$escape_expand( \\
        cmd /c xcopy /Y /D \"$$replace(QT_INSTALL_BINS, /, \\)\\*.dll\" \"$$replace(DESTDIR, /, \)\\) 
    )
}

# macOS特定配置
macx {
    QMAKE_INFO_PLIST = Info.plist  # 可选plist文件配置
    ICON = aviation.icns           # 应用图标
}

# 调试/发布配置
CONFIG(debug, debug|release) {
    DEFINES += _DEBUG
    QMAKE_CXXFLAGS += -g
} else {
    DEFINES += NDEBUG
    QMAKE_CXXFLAGS_RELEASE += -O2
}

# 安装配置（可选）
target.path = $$[QT_INSTALL_BINS]
INSTALLS += target