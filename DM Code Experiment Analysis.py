import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

plt.style.use('ggplot')
# config--处理绘s图中文字符乱码的问题
plt.rcParams['font.sans-serif'] = ['FangSong']   # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False    # 用来正常显示负号

# config--设置文件路径
dirName = r"F:\7 Python\1 Project Python\1 new fuel code size experiment analysis\\"
fileName = r"3.2 新组件编码识别装试验记录 - 控制棒(外径100)  编码位置偏移试验( 尺寸8-12 深度7级) - （读码器倾角15度）- 2018.05.22.xlsx"
EXCEL_PATH = dirName + fileName
COLOR = ['black', 'blue', 'green', 'red', 'gray']	# 定义各表单生成图形的颜色

# 读取Excel各表格的试验数据至df,并绘图
with pd.ExcelFile(EXCEL_PATH) as xls:
    # 遍历excel文件中的每个表单，读取数据并绘制图形
    for i in range(len(xls.sheet_names)):
        sheetName = xls.sheet_names[i]
        df = pd.read_excel(xls, sheetName, header=[1, 2, 3])   # 读取excel数据（指定excel文件中的2、3、4行为标题行）
        codeDeep = df.columns.levels[0]    		# 获取编码深度标题
        codeSize = df.columns.levels[1]    		# 获取编码尺寸标题
        deCodeTime = df.columns.levels[2][2]    # 获取解码时间标题
        # 建立数据分布图及平均值图
        axRowNum = len(codeDeep)	# 数据分布图行数
        axColNum = len(codeSize)	# 数据分布图列数
        label = []		# 平均值图标签
        dataMean = []	# 平均值图数据
        figMean, axMean = plt.subplots(figsize=(8, 8))        # 建立平均值图
        if axRowNum > 1 and axColNum > 1:
            fig, ax = plt.subplots(axRowNum, axColNum, figsize=(16, 8))     # 建立数据分布图
            
            for m in range(axRowNum):
                for n in range(axColNum):
                   
                    ax[m][n].hist(df[codeDeep[m]][codeSize[n]][deCodeTime], color=COLOR[i])
                    ax[m][n].set_ylabel('数量', fontsize=12)
                    ax[m][n].set_xlabel(deCodeTime, fontsize=12)
                    ax[m][n].set_title(sheetName + '\n' + codeDeep[m] + ' ' + codeSize[n], fontsize=12, y = 1.0)
                    label.append(codeDeep[m] + ' ' + codeSize[n])
                    dataMean.append(df[codeDeep[m]][codeSize[n]][deCodeTime].mean())       # 计算平均值,并添加到数组
        else:
            if axRowNum == 1:
                fig, ax = plt.subplots(axRowNum, axColNum, figsize=(16, 8))     # 建立数据分布图
            else:
                fig, ax = plt.subplots(axColNum, axRowNum, figsize=(16, 4))     # 建立数据分布图
				
            for m in range(axRowNum):
                for n in range(axColNum):
                    if axRowNum == 1:
                        x = n
                    else:
                        x = m
                    ax[x].hist(df[codeDeep[m]][codeSize[n]][deCodeTime], color=COLOR[i])
                    ax[x].set_ylabel('数量', fontsize=10)
                    ax[x].set_xlabel(deCodeTime, fontsize=10)
                    ax[x].set_title(codeDeep[m] + '\n' + codeSize[n], fontsize=10, y=1.0)
                    label.append(codeDeep[m] + ' ' + codeSize[n])
                    dataMean.append(df[codeDeep[m]][codeSize[n]][deCodeTime].mean())      # 计算平均值,并添加到数组
            
        # 1显示数据分布图
        fig.tight_layout()
        fig.show()

        # 2配置并显示平均值图
        x = np.arange(len(label))
        barWidth = 0.8
        axMean.bar(x, dataMean, barWidth, color=COLOR[i])
        axMean.set_xticks(x+0.4)
        axMean.set_xticklabels(label, rotation=-70, fontsize=12)
        axMean.set_title(sheetName)
        figMean.tight_layout()
        figMean.show()
