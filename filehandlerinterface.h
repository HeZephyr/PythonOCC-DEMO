#ifndef FILEHANDLERINTERFACE_H
#define FILEHANDLERINTERFACE_H

#include <QString>
#include <QStringList>

// 文件处理器接口 - 插件架构的核心
class FileHandlerInterface {
public:
    virtual ~FileHandlerInterface() {}
    
    // 返回此处理器支持的文件扩展名列表
    virtual QStringList supportedExtensions() const = 0;
    
    // 返回此处理器的描述
    virtual QString description() const = 0;
    
    // 处理指定文件
    virtual bool processFile(const QString &filePath, QStringList &arguments) = 0;
    
    // 获取外部程序路径（如适用）
    virtual QString executablePath() const = 0;
    
    // 是否为外部程序
    virtual bool isExternalProcess() const = 0;
};

// 用于注册文件处理器的宏
#define REGISTER_FILE_HANDLER(className) \
    static FileHandlerRegistration<className> registration;

// 文件处理器注册帮助类
template<class T>
class FileHandlerRegistration {
public:
    FileHandlerRegistration() {
        // 实际的注册逻辑会在FileVisualizerCore中实现
    }
};

#endif // FILEHANDLERINTERFACE_H