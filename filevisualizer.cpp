#include "filevisualizer.h"
#include "configmanager.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QPushButton>
#include <QLabel>
#include <QProgressBar>
#include <QFileDialog>
#include <QMessageBox>
#include <QStatusBar>
#include <QTimer>
#include <QDebug>
#include <QMenu>
#include <QMenuBar>
#include <QApplication>
#include <QFileInfo>
#include <QTime> // Replace QRandomGenerator

FileVisualizer::FileVisualizer(QWidget *parent)
    : QMainWindow(parent),
      m_process(nullptr),
      m_progressTimer(nullptr),
      m_currentProgress(0),
      m_isProcessing(false)
{
    // Initialize random seed for qrand()
    qsrand(QTime::currentTime().msec());
    
    // 初始化配置管理器
    m_configManager = new ConfigManager(this);
    if (!m_configManager->initialize()) {
        QMessageBox::critical(this, tr("配置错误"), 
                             tr("无法加载配置文件。请检查config.ini是否存在且格式正确。"));
        QTimer::singleShot(0, this, &FileVisualizer::close);
        return;
    }
    
    // 初始化UI
    initUI();
    setupMenus();
    setupConnections();
    
    // 创建进程对象
    m_process = new QProcess(this);
    connect(m_process, &QProcess::errorOccurred, this, &FileVisualizer::handleProcessError);
    connect(m_process, QOverload<int, QProcess::ExitStatus>::of(&QProcess::finished), 
            this, &FileVisualizer::handleProcessFinished);
    connect(m_process, &QProcess::readyReadStandardOutput, this, &FileVisualizer::handleProcessOutput);
    
    // 创建进度定时器
    m_progressTimer = new QTimer(this);
    connect(m_progressTimer, &QTimer::timeout, this, &FileVisualizer::updateProgress);
    
    // 更新界面状态
    m_btnCancel->setEnabled(false);
    m_progressBar->hide();
}

FileVisualizer::~FileVisualizer()
{
    if (m_process && m_process->state() != QProcess::NotRunning) {
        m_process->kill();
        m_process->waitForFinished(1000);
    }
}

void FileVisualizer::initUI()
{
    setWindowTitle(tr("文件可视化工具"));
    setMinimumSize(500, 300);
    
    // 创建中央部件
    QWidget *centralWidget = new QWidget(this);
    setCentralWidget(centralWidget);
    
    // 主布局
    QVBoxLayout *mainLayout = new QVBoxLayout(centralWidget);
    
    // 文件选择区
    QHBoxLayout *fileLayout = new QHBoxLayout();
    m_btnSelect = new QPushButton(tr("选择文件"), this);
    m_btnCancel = new QPushButton(tr("取消"), this);
    fileLayout->addWidget(m_btnSelect);
    fileLayout->addWidget(m_btnCancel);
    mainLayout->addLayout(fileLayout);
    
    // 状态显示区
    m_lblStatus = new QLabel(tr("请选择文件进行处理"), this);
    m_lblStatus->setAlignment(Qt::AlignCenter);
    mainLayout->addWidget(m_lblStatus);
    
    // 进度条
    m_progressBar = new QProgressBar(this);
    m_progressBar->setRange(0, 100);
    m_progressBar->setValue(0);
    mainLayout->addWidget(m_progressBar);
    
    // 状态栏
    m_statusBar = new QStatusBar(this);
    setStatusBar(m_statusBar);
    
    // 构建支持的文件扩展名列表，用于状态栏显示
    QStringList allExtensions;
    foreach (const QString &fileType, m_configManager->getSupportedFileTypes()) {
        allExtensions.append(m_configManager->getFileExtensionsForType(fileType));
    }
    
    QString supportedTypes = allExtensions.join(", ");
    m_statusBar->showMessage(tr("支持的文件类型: %1").arg(supportedTypes));
    
    // 添加弹性空间
    mainLayout->addStretch(1);
}

void FileVisualizer::setupMenus()
{
    // 文件菜单
    QMenu *fileMenu = menuBar()->addMenu(tr("文件(&F)"));
    
    QAction *openAction = fileMenu->addAction(tr("打开文件(&O)..."));
    openAction->setShortcut(QKeySequence::Open);
    connect(openAction, &QAction::triggered, this, &FileVisualizer::selectFile);
    
    fileMenu->addSeparator();
    
    QAction *exitAction = fileMenu->addAction(tr("退出(&Q)"));
    exitAction->setShortcut(QKeySequence::Quit);
    connect(exitAction, &QAction::triggered, this, &QWidget::close);
    
    // 帮助菜单
    QMenu *helpMenu = menuBar()->addMenu(tr("帮助(&H)"));
    
    QAction *aboutAction = helpMenu->addAction(tr("关于(&A)"));
    connect(aboutAction, &QAction::triggered, [this]() {
        QMessageBox::about(this, tr("关于文件可视化工具"),
                          tr("<h3>文件可视化工具</h3>"
                             "<p>一个可扩展的文件导入与可视化系统</p>"
                             "<p>版本 1.0</p>"));
    });
}

void FileVisualizer::setupConnections()
{
    connect(m_btnSelect, &QPushButton::clicked, this, &FileVisualizer::selectFile);
    connect(m_btnCancel, &QPushButton::clicked, this, &FileVisualizer::cancelProcessing);
}

void FileVisualizer::selectFile()
{
    // 构建所有支持的文件扩展名
    QStringList allExtensions;
    foreach (const QString &fileType, m_configManager->getSupportedFileTypes()) {
        allExtensions.append(m_configManager->getFileExtensionsForType(fileType));
    }
    
    // 构建文件过滤器
    QString filterText;
    if (!allExtensions.isEmpty()) {
        QStringList formatExtensions;
        foreach (const QString &ext, allExtensions) {
            formatExtensions << "*." + ext;
        }
        filterText = tr("支持的文件 (%1)").arg(formatExtensions.join(" "));
        filterText += ";;" + tr("所有文件 (*)");
    } else {
        filterText = tr("所有文件 (*)");
    }
    
    // 显示文件对话框
    QString fileName = QFileDialog::getOpenFileName(
        this,
        tr("选择要处理的文件"),
        QString(),
        filterText
    );
    
    if (!fileName.isEmpty()) {
        importFile(fileName);
    }
}

bool FileVisualizer::importFile(const QString &filePath)
{
    // 检查文件是否存在
    QFileInfo fileInfo(filePath);
    if (!fileInfo.exists() || !fileInfo.isReadable()) {
        QMessageBox::warning(this, tr("文件错误"), 
                            tr("无法访问文件: %1").arg(filePath));
        return false;
    }
    
    // 自动检测文件类型
    QString fileType = detectFileType(filePath);
    
    if (fileType.isEmpty()) {
        QMessageBox::warning(this, tr("不支持的文件类型"), 
                           tr("不支持处理文件: %1\n扩展名: %2").arg(
                               QFileInfo(filePath).fileName(),
                               QFileInfo(filePath).suffix()));
        return false;
    }
    
    m_currentFilePath = filePath;
    m_currentFileType = fileType;
    
    // 开始处理文件
    return startFileProcessing(filePath, fileType);
}

QString FileVisualizer::detectFileType(const QString &filePath)
{
    // 获取文件扩展名
    QString extension = QFileInfo(filePath).suffix().toLower();
    
    // 查找对应的文件类型
    return m_configManager->getFileTypeForExtension(extension);
}

bool FileVisualizer::startFileProcessing(const QString &filePath, const QString &fileType)
{
    // 如果已经在处理中，先停止
    if (m_isProcessing) {
        cancelProcessing();
    }
    
    // 获取处理程序路径
    QString executablePath = m_configManager->getExecutablePathForType(fileType);
    if (executablePath.isEmpty()) {
        QMessageBox::critical(this, tr("配置错误"), 
                             tr("找不到文件类型 %1 的处理程序").arg(fileType));
        return false;
    }
    
    // 更新UI状态
    QString displayName = m_configManager->getDisplayNameForType(fileType);
    m_lblStatus->setText(tr("正在处理%1文件: %2").arg(
        displayName,
        QFileInfo(filePath).fileName()
    ));
    m_btnSelect->setEnabled(false);
    m_btnCancel->setEnabled(true);
    m_progressBar->setValue(0);
    m_progressBar->show();
    m_isProcessing = true;
    m_currentProgress = 0;
    
    // 启动进程
    QStringList arguments;
    arguments << filePath;
    
    qDebug() << "启动外部程序:" << executablePath << arguments;
    m_statusBar->showMessage(tr("启动 %1 处理程序...").arg(displayName));
    
    m_process->start(executablePath, arguments);
    
    if (!m_process->waitForStarted(5000)) {
        QMessageBox::critical(this, tr("进程错误"), 
                             tr("无法启动处理程序: %1").arg(m_process->errorString()));
        
        // 恢复UI状态
        m_btnSelect->setEnabled(true);
        m_btnCancel->setEnabled(false);
        m_progressBar->hide();
        m_lblStatus->setText(tr("处理失败"));
        m_statusBar->showMessage(tr("处理失败: 无法启动程序"));
        m_isProcessing = false;
        
        return false;
    }
    
    // 启动进度定时器 (每200毫秒更新一次)
    m_progressTimer->start(200);
    
    // 发送开始处理信号
    emit processingStarted(filePath, fileType);
    
    return true;
}

void FileVisualizer::handleProcessError(QProcess::ProcessError error)
{
    QString errorMessage;
    
    switch (error) {
        case QProcess::FailedToStart:
            errorMessage = tr("无法启动程序, 检查路径和权限");
            break;
        case QProcess::Crashed:
            errorMessage = tr("程序异常崩溃");
            break;
        case QProcess::Timedout:
            errorMessage = tr("程序响应超时");
            break;
        case QProcess::WriteError:
            errorMessage = tr("无法向程序写入数据");
            break;
        case QProcess::ReadError:
            errorMessage = tr("无法从程序读取数据");
            break;
        default:
            errorMessage = tr("未知错误");
            break;
    }
    
    qDebug() << "处理程序错误:" << errorMessage;
    
    if (error != QProcess::Crashed || m_progressTimer->isActive()) {
        // 停止进度定时器
        m_progressTimer->stop();
        
        // 恢复UI状态
        m_btnSelect->setEnabled(true);
        m_btnCancel->setEnabled(false);
        m_isProcessing = false;
        
        m_lblStatus->setText(tr("处理失败: %1").arg(errorMessage));
        m_statusBar->showMessage(tr("处理失败: %1").arg(errorMessage));
        
        // 仅在非用户取消时显示错误对话框
        QMessageBox::critical(this, tr("处理错误"), 
                             tr("文件处理出错: %1").arg(errorMessage));
        
        // 发送处理完成信号(失败)
        emit processingFinished(false, errorMessage);
    }
}

void FileVisualizer::handleProcessFinished(int exitCode, QProcess::ExitStatus exitStatus)
{
    // 停止进度定时器
    m_progressTimer->stop();
    
    // 恢复UI状态
    m_btnSelect->setEnabled(true);
    m_btnCancel->setEnabled(false);
    m_isProcessing = false;
    
    QString resultMessage;
    bool success = false;
    
    if (exitStatus == QProcess::NormalExit && exitCode == 0) {
        m_progressBar->setValue(100);
        QString displayName = m_configManager->getDisplayNameForType(m_currentFileType);
        resultMessage = tr("%1文件处理成功完成").arg(displayName);
        success = true;
    } else {
        resultMessage = tr("处理未成功完成 (退出代码: %1)").arg(exitCode);
    }
    
    m_lblStatus->setText(resultMessage);
    m_statusBar->showMessage(resultMessage);
    
    // 发送处理完成信号
    emit processingFinished(success, resultMessage);
}

void FileVisualizer::handleProcessOutput()
{
    // 读取输出，可用于更新实际进度
    QByteArray output = m_process->readAllStandardOutput();
    QString outputStr = QString::fromUtf8(output).trimmed();
    
    if (!outputStr.isEmpty()) {
        // 可以在这里处理进度输出，如果外部程序提供了进度信息
        // 例如，解析输出中的进度百分比: "Progress: 45%"
        QRegExp progressRegex("Progress:\\s*(\\d+)%");
        if (progressRegex.indexIn(outputStr) != -1) {
            int progress = progressRegex.cap(1).toInt();
            m_currentProgress = progress;
            m_progressBar->setValue(progress);
            
            // 发送进度信号
            emit processingProgress(progress);
        }
        
        // 输出可能包含状态信息，可以更新到状态栏
        m_statusBar->showMessage(outputStr);
    }
}

void FileVisualizer::updateProgress()
{
    // 如果没有明确的进度，就模拟进度
    if (m_isProcessing) {
        // 使用渐进式进度模拟
        if (m_currentProgress < 90) {
            // 使用qrand()替代QRandomGenerator
            m_currentProgress += 1 + (qrand() % 3); // 随机增加1-3%
            if (m_currentProgress > 90) {
                m_currentProgress = 90; // 最高到90%，剩下10%在完成时更新
            }
            m_progressBar->setValue(m_currentProgress);
            
            // 发送进度信号
            emit processingProgress(m_currentProgress);
        }
    }
}

void FileVisualizer::cancelProcessing()
{
    if (!m_isProcessing) return;
    
    if (QMessageBox::question(this, tr("确认取消"),
                            tr("确定要取消当前处理任务吗?"),
                            QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        
        if (m_process->state() != QProcess::NotRunning) {
            // 首先尝试正常终止
            m_process->terminate();
            
            // 给一些时间让进程自己退出
            if (!m_process->waitForFinished(2000)) {
                m_process->kill(); // 强制终止
            }
        }
        
        // 停止进度定时器
        m_progressTimer->stop();
        
        // 恢复UI状态
        m_btnSelect->setEnabled(true);
        m_btnCancel->setEnabled(false);
        m_progressBar->setValue(0);
        m_isProcessing = false;
        
        m_lblStatus->setText(tr("处理已取消"));
        m_statusBar->showMessage(tr("处理已取消"));
        
        // 发送处理完成信号(取消)
        emit processingFinished(false, "处理已取消");
    }
}