#include <QApplication>
#include "filevisualizer.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);
    
    // 设置应用信息
    QCoreApplication::setOrganizationName("COMAC");
    QCoreApplication::setApplicationName("FileVisualizer");
    
    FileVisualizer visualizer;
    visualizer.show();
    
    return app.exec();
}