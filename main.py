import sys
from PyQt5.QtWidgets import QApplication
from app_ui import EcommerceAnalysisApp
from matplotlib import rcParams

# 设置中文字体
rcParams['font.sans-serif'] = ['Microsoft YaHei'] # 设置中文字体为微软雅黑
rcParams['axes.unicode_minus'] = False # 解决负号显示问题

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格使界面更现代化
    
    # 设置应用程序字体
    from PyQt5.QtGui import QFont
    font = QFont("Microsoft YaHei", 9)  # 使用微软雅黑字体
    app.setFont(font)
    
    # 创建并显示主窗口
    window = EcommerceAnalysisApp()
    window.show()
    
    sys.exit(app.exec_()) 