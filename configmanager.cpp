#include "configmanager.h"
#include <QSettings>
#include <QFileInfo>
#include <QDir>
#include <QDebug>

ConfigManager::ConfigManager(QObject *parent) 
    : QObject(parent)
{
}

bool ConfigManager::initialize(const QString &configPath)
{
    // 使用指定配置或默认配置
    m_configPath = configPath.isEmpty() ? QDir::currentPath() + "/../config.ini" : configPath;
    
    QFileInfo configFile(m_configPath);
    if (!configFile.exists() || !configFile.isReadable()) {
        qCritical() << "配置文件不存在或不可读:" << m_configPath;
        return false;
    }
    
    QSettings settings(m_configPath, QSettings::IniFormat);
    
    // 读取所有支持的文件类型
    settings.beginGroup("FileTypes");
    QStringList fileTypes = settings.childGroups();
    settings.endGroup();
    
    if (fileTypes.isEmpty()) {
        qCritical() << "配置文件中未定义任何文件类型";
        return false;
    }
    
    // 读取每种文件类型的配置
    foreach (const QString &fileType, fileTypes) {
        QString groupPath = QString("FileTypes/%1").arg(fileType);
        settings.beginGroup(groupPath);
        
        // 读取可执行文件路径
        QString exePath = settings.value("executablePath").toString().trimmed();
        
        // 检查可执行文件路径
        QFileInfo exeFile(exePath);
        if (exePath.isEmpty() || !exeFile.exists() || !exeFile.isExecutable()) {
            qWarning() << "文件类型" << fileType << "的可执行文件路径无效:" << exePath;
            settings.endGroup();
            continue;
        }
        
        m_executablePaths[fileType] = exePath;
        
        // 读取文件扩展名列表
        QStringList extensions = settings.value("extensions").toString().split(",", QString::SkipEmptyParts);
        for (int i = 0; i < extensions.size(); ++i) {
            extensions[i] = extensions[i].trimmed().toLower();
            // 创建扩展名到类型的映射
            m_extensionToType[extensions[i]] = fileType;
        }
        
        if (extensions.isEmpty()) {
            qWarning() << "文件类型" << fileType << "未定义扩展名";
        }
        
        m_fileExtensions[fileType] = extensions;
        
        // 读取显示名称
        QString displayName = settings.value("displayName").toString().trimmed();
        if (displayName.isEmpty()) {
            displayName = fileType;
        }
        
        m_displayNames[fileType] = displayName;
        
        settings.endGroup();
    }
    
    return validateConfiguration();
}

bool ConfigManager::validateConfiguration()
{
    if (m_executablePaths.isEmpty()) {
        qCritical() << "没有配置有效的文件处理程序";
        return false;
    }
    
    return true;
}

QString ConfigManager::getExecutablePathForType(const QString &fileType) const
{
    return m_executablePaths.value(fileType);
}

QStringList ConfigManager::getSupportedFileTypes() const
{
    return m_executablePaths.keys();
}

QStringList ConfigManager::getFileExtensionsForType(const QString &fileType) const
{
    return m_fileExtensions.value(fileType);
}

QString ConfigManager::getFileTypeForExtension(const QString &extension) const
{
    return m_extensionToType.value(extension.toLower());
}

QString ConfigManager::getDisplayNameForType(const QString &fileType) const
{
    return m_displayNames.value(fileType, fileType);
}

void ConfigManager::setExecutablePathForType(const QString &fileType, const QString &path)
{
    m_executablePaths[fileType] = path;
    
    // 保存到配置文件
    QSettings settings(m_configPath, QSettings::IniFormat);
    settings.setValue(QString("FileTypes/%1/executablePath").arg(fileType), path);
}