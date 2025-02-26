#include <QApplication>
#include <QWidget>
#include <QPushButton>
#include <QFileDialog>
#include <QVBoxLayout>
#include <QLabel>
#include <QProcess>
#include <QMessageBox>
#include <QDebug>
#include <QSettings>
#include <QFileInfo>
#include <QFile>

class MainWindow : public QWidget {
    Q_OBJECT
public:
    MainWindow(QWidget* parent = nullptr) : QWidget(parent) {
        initUI();
        // 读取外部配置文件中的路径（若配置项不存在则报错退出）
        readConfiguration();

        // QProcess 对象由 MainWindow 统一管理
        m_process = new QProcess(this);
        connect(m_process, &QProcess::errorOccurred,
                this, &MainWindow::onProcessError);
        connect(m_process,
                static_cast<void(QProcess::*)(int, QProcess::ExitStatus)>(&QProcess::finished),
                this, &MainWindow::onProcessFinished);
    }

private slots:
    // 选择文件槽
    void selectFile() {
        QString fileName = QFileDialog::getOpenFileName(
            this,
            "选择 XML 文件",
            "",
            "XML 文件 (*.xml)"
        );
        if (!fileName.isEmpty()) {
            m_lblStatus->setText(QString("已选择文件: %1").arg(fileName));

            // 如果之前的外部程序还在运行，先将其杀掉
            if (m_process->state() != QProcess::NotRunning) {
                qDebug() << "杀掉之前运行的程序";
                m_process->kill();
                m_process->waitForFinished(3000); // 最多等待 3 秒退出
            }

            // 通过配置项获取外部程序路径
            qDebug() << "启动外部程序：" << m_executablePath;
            m_process->start(m_executablePath, QStringList() << fileName);

            // 等待启动阶段，若启动失败则弹出错误提示
            if (!m_process->waitForStarted()) {
                QMessageBox::critical(this, "错误",
                                      QString("无法启动外部程序: %1")
                                      .arg(m_process->errorString()));
                m_btnSelect->setEnabled(true);
            }
        }
    }

    // 外部程序出错槽：只对启动失败弹出错误提示，其它错误仅记录日志
    void onProcessError(QProcess::ProcessError error) {
        if (error == QProcess::FailedToStart) {
            QMessageBox::critical(this, "错误",
                                  QString("无法启动外部程序: %1")
                                  .arg(m_process->errorString()));
        } else {
            qDebug() << "外部程序运行错误:" << m_process->errorString();
        }
    }

    // 外部程序结束时更新状态
    void onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus) {
        qDebug() << "外部程序退出, exitCode:" << exitCode
                 << ", exitStatus:" << exitStatus;
        m_lblStatus->setText("外部程序运行结束");
    }

private:
    // 从外部配置文件中读取可执行文件路径，若配置文件或配置项不存在则报错退出
    void readConfiguration() {
        QString configPath = "../config.ini";
        if (!QFileInfo::exists(configPath)) {
            QMessageBox::critical(this, "配置错误",
                                  QString("配置文件 [%1] 不存在，请检查配置！").arg(configPath));
            qCritical() << "配置文件" << configPath << "不存在";
            exit(1);
        }

        QSettings settings(configPath, QSettings::IniFormat);
        QVariant value = settings.value("Paths/executablePath");
        if (!value.isValid() || value.toString().trimmed().isEmpty()) {
            QMessageBox::critical(this, "配置错误",
                                  "未配置可执行文件路径，请检查配置文件中的 [Paths/executablePath] 项！");
            qCritical() << "未配置可执行文件路径";
            exit(1);
        }

        m_executablePath = value.toString().trimmed();
        qDebug() << "配置文件中可执行文件路径:" << m_executablePath;

        QFileInfo exeInfo(m_executablePath);
        if (!exeInfo.exists() || !exeInfo.isFile()) {
            QMessageBox::critical(this, "配置错误",
                                QString("配置的可执行文件 [%1] 不存在或不是文件，请检查配置！")
                                .arg(m_executablePath));
            qCritical() << "配置的可执行文件" << m_executablePath << "不存在或不是文件";
            exit(1);
        }
    }

    void initUI() {
        setWindowTitle("布线系统 demo");
        setGeometry(100, 100, 400, 200);

        QVBoxLayout* layout = new QVBoxLayout(this);
        m_btnSelect = new QPushButton("选择 XML 文件", this);
        connect(m_btnSelect, &QPushButton::clicked, this, &MainWindow::selectFile);
        layout->addWidget(m_btnSelect);

        m_lblStatus = new QLabel("当前未选择文件", this);
        layout->addWidget(m_lblStatus);

        setLayout(layout);
    }

private:
    QPushButton* m_btnSelect;
    QLabel* m_lblStatus;
    QProcess* m_process;
    QString m_executablePath;   // 外部程序路径，从配置文件中读取
};

#include "main.moc"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);
    MainWindow window;
    window.show();
    return app.exec();
}