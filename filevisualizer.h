#ifndef FILEVISUALIZER_H
#define FILEVISUALIZER_H

#include <QMainWindow>
#include <QProcess>
#include <QMap>
#include <QFileInfo>

QT_BEGIN_NAMESPACE
class QLabel;
class QPushButton;
class QProgressBar;
class QStatusBar;
class QTimer;
QT_END_NAMESPACE

class ConfigManager;

class FileVisualizer : public QMainWindow {
    Q_OBJECT

public:
    explicit FileVisualizer(QWidget *parent = nullptr);
    ~FileVisualizer();

    // 公共API，允许从外部调用导入文件功能
    bool importFile(const QString &filePath);
    
signals:
    // 可用于外部系统监听进度
    void processingStarted(const QString &filePath, const QString &detectedType);
    void processingProgress(int percentage);
    void processingFinished(bool success, const QString &message);

private slots:
    void selectFile();
    void handleProcessError(QProcess::ProcessError error);
    void handleProcessFinished(int exitCode, QProcess::ExitStatus exitStatus);
    void handleProcessOutput();
    void updateProgress();
    void cancelProcessing();

private:
    void initUI();
    void setupMenus();
    void setupConnections();
    bool startFileProcessing(const QString &filePath, const QString &fileType);
    QString detectFileType(const QString &filePath);
    
    // UI组件
    QPushButton *m_btnSelect;
    QPushButton *m_btnCancel;
    QLabel *m_lblStatus;
    QProgressBar *m_progressBar;
    QStatusBar *m_statusBar;
    
    // 核心功能组件
    ConfigManager *m_configManager;
    QProcess *m_process;
    QTimer *m_progressTimer;
    
    // 状态变量
    QString m_currentFilePath;
    QString m_currentFileType;
    int m_currentProgress;
    bool m_isProcessing;
};

#endif // FILEVISUALIZER_H