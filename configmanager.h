#ifndef CONFIGMANAGER_H
#define CONFIGMANAGER_H

#include <QObject>
#include <QMap>

class ConfigManager : public QObject {
    Q_OBJECT
    
public:
    explicit ConfigManager(QObject *parent = nullptr);
    
    bool initialize(const QString &configPath = QString());
    
    QString getExecutablePathForType(const QString &fileType) const;
    QStringList getSupportedFileTypes() const;
    QStringList getFileExtensionsForType(const QString &fileType) const;
    QString getFileTypeForExtension(const QString &extension) const;
    QString getDisplayNameForType(const QString &fileType) const;
    
    // 用于从外部设置配置的API
    void setExecutablePathForType(const QString &fileType, const QString &path);
    
private:
    bool validateConfiguration();
    
    QString m_configPath;
    QMap<QString, QString> m_executablePaths;       // 类型 -> 可执行文件路径
    QMap<QString, QStringList> m_fileExtensions;    // 类型 -> 扩展名列表
    QMap<QString, QString> m_extensionToType;       // 扩展名 -> 类型
    QMap<QString, QString> m_displayNames;          // 类型 -> 显示名称
};

#endif // CONFIGMANAGER_H