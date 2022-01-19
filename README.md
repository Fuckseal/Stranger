


##### 1. 设置对象属性 

对象名.属性名=属性值
**例如**
	Label1.Caption =Text1.Text & ", 您好！欢迎光临！"
##### 2. 调用对象的方法
	对象名.方法名 [参数列表]
	
##### 3.定义常量
`[Public|Private]Const 常量名 [AS 类型] = 常量表达式`

 **例如**
	` Const MyWar =359  ‘默认私有常量
	   Const MyStr = “Hello”， MyDouble as Double = 3.456 ’在一行声明多个常量`
>在Visual Basic 中，可以使用Print 方法来**显示**文本和数据
	
`[对象. ] Print  [表达式]  [,|;]  [,|;] ...`
其中”对象“可以是**Form、PictureBox、Printer、Debug**
>在使用Print方法显示文本时，可以通过选项Spc（n）和Tab（n）来控制字符的位置
##### Cls方法
用于**清除**程序运行时窗体或图片框所生成的**图形和文本。**
`[对象] .Cls`
>调用Cls方法之后,对象的下一次**打印**坐标CurrentX 和 CurrentY 属性将**复位**为0.
##### 声明变量
在Visual Basic 中，可用Dim语句来**声明变量**的数据类型并分配内存空间
`Dim|Private|Static|Public <变量名> [AS类型]    [，<变量名>  [AS 类型] ]`
##### 赋值语句 
将表达式的值**赋给**变量或对象的属性，可使用赋值语句来实现
`[Let] 变量名 | 属性名 = 表达式`
如果要把**多个**赋值语句放在同一行，各个语句之间必须用冒号**隔开**
`a=3 : b=4 : c=5`
注释语句是为了方便程序**阅读**对程序进行的说明，对程序运行没有影响
`Rem 注释语句
‘ 注释语句`
##### 结束语句
`End
Unload Me <对象名称>`
End语句用于**结束**正在运行的程序，提供了一种**强迫中止**程序的方法
Unload语句用于从内存中**卸载**窗体或控件。卸载床提前，依次发生窗体的**QueryUnload 和Unload **事件过程。
#### If 语句

**单行格式**
`If 条件 Then 语句1 [Else 语句2]
多行块格式
If <条件1> Then
	语句块 1
ElseIf <条件2>  Then
	语句块2
ElseIf <条件3> Then
	语句块3
	...
Else
	语句块n
End If`
##### IIf函数
用于**计算**表达式的值并据此**返回**两个值中的一个。
`Resoult = IIF（条件，True部分，False部分）`
##### Select Case语句
`Select Case <此时表达式>
Case 表达式列表1
	语句块1
Case 表达式列表2
	语句块2
		...
Case  Else
	 语句块n
End Select`
在每个Case子句中可以使用**多重**表达式或使用范围，例如
`Case 1 To 4，7 To 9，11，13，Is >MaxNumber`
也可以针对**字符串**指定范围和多重表达式。
`Case “everything”，“buts” To “soup”，TestItem`
##### For 循环
For 循环用于**重复**执行若干语句
`For 循环变量=初值 To 终值 [Step 步长]
	语句块
	[Exit For]
	语句块
Next [循环变量]`
##### While循环
根据指定条件**重复**执行一个或多个语句
`While 条件
	语句块
Wend`
##### 前测型Do循环语句
`Do While | Until < 循环条件>
	[语句块]
	[Exit Do]
	[语句块]
Loop`
##### 后测型Do循环语句
`Do
	[语句块]
	[Exit Do]
	[语句块]
Loop While| Until  <循环条件>`
##### 定长数组
`Dim 数组名（[下标下界 To]下标上界 [，下标下届 To 下标上界]）[AS 数据类型]`
- 若生略AS子句，则定义的数组为***Variant***类型
- **数值数组**全部元素**初始化**为0，**字符串数组**中的全部元素初始化为**空字符串**
- 下标可以是不超过**Long数据类型**的范围（-2，147，483，648 到 2...647）的整数
- 如果省略"下标下届To"，则数组默认下界为0.假如希望下标从1 开始，则应通过Option Base1设置
- 数组的维数最多可以有**60维**，下标只能是**常数**，不能是变量和表达式
##### LBound和UBound

**LBound 函数**返回数组某一维的**下界值**，而UBound函数返回数组某一维的**上界值**这两个函数一起使用可以确定一个数组的大小

`LBound（数组名[，维数]）
UBound（数组名[，维数]）`
##### 数组初始化
- 对于**Variant**类型数组，可以使用Visual Basic提供的**Array函数**进行**初始化**
`数组名 = Array（值列表）`
- 数组变量只能能是Variant类型，只适用于**一维**数组。默认Array函数创建的数组下界**从0**开始。
##### 生成随机数
使用Rnd（）函数可以**生成**一个包含随机数值的**单精度浮点**数，其值在[0，1）

>**生成** 1 ~ 100 范围内的随机整数  ：Int（101 * Rnd）
	ReDim 语句用来**声明**或重新声明原来用Privatre、Public或Dim语句声明过的二带**空圆括号**的**动态**数组的大小

`ReDim [Preserve] 变量 (下标，下标) AS 数据类型`
>ReDim语句只能出现在**事件**过程或**通用**过程中，因为它定义的数组是个**临时**数组。

**数组的清除**

`Erase 数组名 [,数组名] ...`
如果是**Variant**数组，每个元素置为**Empty**
如果是**对象**数组，则将每个元素设为**Nothing**
通用过程具有作用**范围、名称、参数列表**和**过程体**可以使用Sub语句来声明
`[ Private | Public  ]    [ Static ]  Sub 过程名   [（参数列表）]
	[语句块]
	[Exit Sub]
	[语句块]
End Sub`
>如果没有**显式指定**Public、Private关键字，则Sub过程**默认**范围是**Public**

##### 窗体事件过程
`Private Sub Form_事件名  [(参数列表)]
	语句块
End Sub`

**控件事件过程**

`Private Sub 控件名_事件名 [(参数列表)]
	语句块
End Sub`
声明过程时，参数应遵循的格式
`[Optional]  [ByVal | ByRef]   [ParamArray]  变量名[（）]   [AS 数据类型]`
	>Optional表示参数**可选**，
	 ByVal表示**按值**传递，
      ByRef表示按**地址**传递
      ParamArray表示参数数目任意，用于声明**动态**数组

例如
`Sub OpDemo（a As Integer，Optional b As Integer）
	print a，b
End Sub
Private Sub Form_Click（）
	OpDemo 10，20
	OpDemo 10 
End Sub
运行结果为
	10      20
	10      0`
声明时可以为**可选**参数指定默认值
`Sub OpDemo（a As Integer，Optional b As Integer = 1 ）`
>ParamArray 关键字**只能**用于参数列表中**最后一个参数**，
>且**不能**与*ByVal、ByRef、Optional*关键字**一起**使用

例如
`Sub PaDemo（ParamArray a（））`
##### Call语句
Call语句**调用**过程
`Call 过程名 [(实际参数)]
`
也可以把过程名作为一个**语句**使用
`过程名 [实际参数]
`
##### Rnd语句
>使用随机函数Rnd语句时可以在之前加一条
随机数生成器**初始化**语句***Randmize***

`Randmize [数值]
`
生成某个范围**随机整数**可以使用此表达式
`Int（（上限 - 下线 + 1 ）* Rnd(  )  + 1）
数学表达式sin45° + cos45° + log24可以写成
sin（45*（3.1415926/180））+cos（45*（3.1415926/180））+log（4）/（2）
 `
##### 格式化输出函数
`Format（表达式[，格式化字符串]）

Dim a As Double

a=1234.567
Print Format（a,"00000.0000"）,Format（a,"##,##.####"）
Print Format（a,"#####.##%"）,Format（a,"$#####.##"）
Print Format（a,"+#####.##"）,Format（a,"0.0000E+00"）
Print Format（a,"yyyy年mm月dd日 hh点mm分ss秒"）

>输出结果
01234.5670  1, 234.567
123456.7%   $1234.567
+1234.567    1.2346E+03
2008年03月16日16点45分36秒
`
##### 自定义函数
即**Function**过程
`[Private | Public ]  [Static]  Function <函数名> ([参数列表])  [As 数据类型]
	[语句块] 
	[函数名=表达式]
	[Exit Function]
	[语句块]
	[函数名=表达式]
End Function
`
>如果**未显示**指定Public、Private关键字，函数默认范围为**Public**
自定义函数调用方法和内部函数调用方法相同


##### 错误捕捉语句On Error

**激活**错误捕捉，并将错误处理程序**指定**为从**行号**位置开始的程序段
`On Error Goto [行号]
`
##### 常用结构
`Sub ErrorDemo（）
	[没有错误的语句块]
	On Error Goto Error ErrorHandler            ’启用错误捕捉功能
	[可能会出现错误的语句块]       
	Exit Sub
ErrorHandler ：
	[错误处理语句]          ’错误处理由此开始
End Sub
`
**Err对象**用于**捕获错误**信息
属性**Number**为当前错误的**编号**，**Description**属性当前错误的**描述**


## 一、标签控件：
#### Label
###### 常用属性：
- Name **名称 **
- Backcolor 文本和图形的**背景**色
- ForColor 文本和图形的**前景**色
- Caption 控件中显示的**文本**
- Enabled **可用**性 Visible **可见**性
- Font **字体**（如宋体、楷体等）
- Fontname **字体名称**
- Fontbold **粗体**
- Fontsize **字号**
- Height **高度** Width **宽度**
- left 控件**左**边缘与容器**左**边缘的**距离**
- top 控件**上**边缘与容器**上**边缘的**距离**
- Alignment （文本的）**对齐**方式
- Autosize **自动改变大小**
- Backstyle **背景样式** Borderstyle **边框样式**
- Wordwrap 是否进行**水平或垂直展开**
#####  2、常用方法:
Move：移动
`语法：Object.Move Left,Top,Width,Heitht）`
3、常用事件：
- Change：**改变**
- Click 鼠标**单击**时 Dblclick 鼠标**双击**时
- MouseUp **按下**鼠标按钮时 MousMove **移动**鼠标时
#### 二、 文本框控件 Textbox
 （无 Caption 属性）
**常用属性：**
- Maxlength：**最大**字符数 Multiline 是否允许**多行** （取值：True / False）
- PasswordChar：用户所输字符字符是否要**显示**出来（值通常为*，用作**密码输入**框）
- Scrollbars：**滚动条** 取值 0：无 1：有水平 2:有垂直 3：水平、垂直都有
- Text：文本框中的**文本** Locked：**锁定**（True 锁定，不可用 False 不锁定）
- SelLength：所选**字符数**
- SelStart：（所选文本的**起始点** 例：第一个字符，值为 0）
- SelText：所选的**文本**
- TabIndex： 访问 Tab 键的**顺序** TabStop：是否可用 Tab 键**选定**
**常用方法：**
- SetFocus：用于将焦点移至文本框控件（即**获得焦点**）
- Move： **移动**
**常用事件：**
- Change：内容改变时发生 **KeyDown、KeyUp、KeyPress：**键盘事件
- LostFocus：**失去焦点**时发生 GotFocus：**获得**焦点时发生
#### 三、 命令按钮控件 CommandButton
1、常用属性：
- Caption：（按钮上）显示的**文本**
- Cancel：是否是**取消按钮**（True：是，可按 ESC 键选中）
- Default：**默认按钮**（True：是，可按 Enter 键选中）
- Style：**样式、类型**（0:标准按钮 1:图像式按钮）
- Picture：（Style 值为 1 时）调用图形文件
- Value：**是否被选中**（True：已选中；False:未选中）
- ToolTipText：**工具提示符**（即光标徘徊时的显示字符串）
2、常用事件：
- Click ：**鼠标单击** （无双击事件）
也有：MouseDown MouseUp KeyDown KeyUp
#### 四、单选按钮控件 OptionButton(一组中只能选择一个)
1、常用属性：
- Alignment 文本的**对齐方式**（0：文本在左 1:文本在右 2：居中）
- Caption：按钮旁边的**提示文本**
- Value 是否**被选中**（True:选中 False：未选中）
2、常用事件：click
#### 五、复选框控件 CheckBox (一组中可同时选多个)
**常用属性：**
- Alignment **对齐方式**（0 1 2：左 右 中）（和单选按钮一样）
- Caption 按钮旁边的**提示文本**
- Value 是否**被选中**（0：未选中； 1：已选中；2：不可用）
**常用事件：** Click
#### 六、框架控件 Frame (容器控件，可对其他控件分组)
**常用属性：**
- Caption 框架左上方显示的**文本**
**常用事件：**
- Click **单击**Dblclick **双击** 
#### 七、滚动条控件（水平：HScrollBar、垂直：VScrollBar）
常用属性：
- LargeChange: 大变化（单击滚动**箭头**时 Value 的改变量）
- SmallChange: 小变化（单击滚动**区域或拖动滑块**时 Value 属性的改变量）
- Max: Value 的最大值（滑块位于最**右端、最下端**时）
- Min：Value 的最小值（滑块位于最**左端、最上端**时）
- Value：滚动框(滑块)的**当前位置**（取值范围：0~32767）
常用事件：
- Change：滚动或通过**代码改变 Value **属性值时发生
- Scroll：（**拖动**引起改变）（故 Scroll 事件**先**发生，Change 事件**后**发生）
#### 八、计时器控件 Timer
**常用属性：**
- Enabled **有效性**（True:有效,启动 False:无效,停止）
- Interval 时间**间隔**，即每隔多久触发一次 Timer 事件
（单位：毫秒 1 秒=1000 毫秒）
**常用事件：**
只有一个**Timer** 事件
#### 九、列表框控件 ListBox
常用属性：
- List 列表框中的所有项目组成的**字符串数组**
- ListCount 项目个数 Columns **栏数、列数**
- ListIndex 项目索引 Selected 一个项的**选择状态**（true、false）
- Sorted 是否按字母表顺序**排序**（默认值：False，按添加的先后顺序排序）
- Style **样式**（0：标准列表框 1：复选框式列表框）
- Text 选中的**文本**（设计时不可用）
###### 常用方法：
- Additem **添加**一项 RemoveItem**删除**一项 Clear **清除**所有项
#### 十、组合框控件 ComboBox
**常用属性：**
- Style 0：**下拉式组合**框（可选可输）
- **简单组合**框（可选可输）（支持 Dblclick 事件）
- **下拉式列表**框（只可选不可输）
- Text：
（1）编辑域中的**文本**（Style 取值为 0、1 时）
（2）**选中的**项目（Style 取值为 2 时
#### VB  项目五  制作多媒体程序  知识汇总
- ==坐标==：描述一个**像素**在屏幕上**的位置**或打印纸上的**点的位置**。
（窗体上的任何一点都可以用X坐标和Y坐标来表示。）
- 窗体==默认坐标系==：**原点**在左上角，**X轴**正方向向右，**Y轴**正方向向下。
- ==ScaleMode==属性：用于设置坐标的度量**单位**，默认度量单位为**缇**。（还可以是：点、英尺、厘米等）

- ==Scale==方法：用于建立**自定义**坐标系。
- VB中坐标的表示方式有两种：**绝对**坐标和**相对**坐标。
绝对坐标：是**相对于原点**的横向距离与纵向距离。
相对坐标：是**相对于最后参照点**的横向距离与纵向距离。
坐标前**有**Step表示**相对**坐标，**没有**Step表示**绝对**坐标。
#### 1、VB中颜色的表示方法有以下三种：
（1）使用系统自带的颜色**常数**，如VBRed、VBGreen等；
（2）使用**QBColor函数**，共可表示16种颜色，参数的取值范围为**0~15**；
（3）使用**RGB**(red,green,blue)**函数**，各参数的取值范围均为**0~255**。
- Pset方法：用于将对象上的**点**设置为**指定颜色**。（简称画点）
- DrawWidth：**线宽**
- Line方法：用于在窗体或图像框中**画直线和矩形**。
- Circle方法：用于在对象上**画圆、椭圆或弧**。
- Line控件：是线条控件，它可以**显示水平线、垂直线或者对角线**。
常用属性：
- AutoRedraw：**自动重绘**   
- BorderColor：**边框颜色**
- BorderStyle：**边框样式**（取值：**0，1~6**：透，实  虚  点  点  双 内收）
  - Borderwidth：**边框厚度**，默认值为**1**。
#### 2、Shape控件：是图形控件，可用于显示矩形、正方形、椭圆、圆、圆角矩形、圆角正方形。
   常用属性：
- Shape：用于设置所**显示的图形**。（0~5：矩  正  椭  圆  矩  正）
- FillColor：**填充颜色**   （只有封闭图形才能填充）
  - FillStyle：**填充样式**（取值：**0~7**， 实 透 水 垂   左 右 交 对）
  （即：实心、透明、水平线、垂直线、左上对角线、右上对角线、交叉线、对角交叉线）
#### 3、图像框控件（PictureBox）   P152-153
**常用属性：**
- AutoReDraw：**自动重绘**（为True：重绘有效  False：无效）
- AutoSize：控件是否**自动改变大小**
（取值为True：控件变   False：控件、图形都不变）
- Height、Width：**高度和宽度**
   - Picture：控件中要**显示的图片**
**常用方法：**
- PaintPicture：**绘制图形**（简称画图），其语法格式为：
`Object.paintPicture  picture,x1,y1,Width1,height1, x2,y2,width2,height2,opcode`
 - LoadPicture：加载图像
         ` 语法格式：Object.Picture=loadPicture([Filename])`
 例如：
	 `picture1.picture=loadPicture("F:\图片\1.jpg")`
>    说明：若**省略**FileName，则**清除**窗体、图像框及图像控件中的图形。

#### 5、APP对象：是通过关键字APP访问的全局对象,通过它指定可获取以下信息： 
应用程序的**标题、版本信息、可执行文件和帮助文件**的**路径及名称** 以及是否
允许前一个应用程序的示例。
     通过APP对象的**Path属性**可返回或设置当前程序所在的路径，该属性**设计时不可用，运行时只读**。
#### 6、图像控件（Image） ：专门用于显示图像。  （√）（P 152）
**常用属性**
- Picture：要**显示的图片      **
- Tag：用来**存储额外数据**
- Stretch：图形是否**调整大小**（True：图形变   False：控件变）
   常用方法：
   - Move 用于**移动图像控件**。
  `   语法格式为：Object.Move  Left,top,width,height  (顺序：左 上 宽 高)`
#### 7、图像框控件与图像控件的异同点：
 - **图形**可以放在**窗体上**，也可以放在**图像框控件或图像控件**上。图像控件专门用来**显示图像**，而窗体和图像框**除了**可以显示图像外，**还**提供了**画图的方法**，可以在运行时画图。（√）
- 图像框控件除了可以**接受和输出**一般图形以外，还可用于**创建动态画图**并支持**Print方法**，因此可以在对象上**输出文本**。而图像控件**只能**用来**显示图像**。（√）
#### 8、Declare语句：用于在模块级别中声明对DLL动态链接库中外部过程的引用。
#### 9、API函数：使用字符串作为操作命令来控制媒体的设置。
常用命令：
（1）Open： **打开**   
（2）Close：**关闭**   
（3）Play：  **播放**
（4）Pause: **暂停**  
（5）Stop： **停止**    
（6）Seek:   设置**播放位置**
（7）Set：   设置**设备状态**                 
（8）Status: 确定**当前状态**
#### 10、ShockWaveFlash控件：Flash动画播放器
常用属性：
- Movie：要播放的**Flash动画文件**
- TotalFrames：**总**帧数   
- CurrentFrame：**当前**帧
**常用方法：**
- Play：**开始**播放   
- Back：跳到**上一帧**（后退）
- Forward：跳到**下一帧**（前进） 
- Rewind：返回**第一帧**
- Stop：**暂停**播放
#### 11、WindowsMediaPlayer：媒体播放器
###### 常用属性：
- URL：**文件位置**
- EnableContextMenu：是否显示**右键菜单**
- FullScreen：是否**全屏显示**
- StretchToFit：是否**伸展到最佳大小**
- uiMode：设置**播放模式**（取值为Full：包含控制条   None：没有控制条）
- PlayState:**当前状态**（1：已停止   2：暂停   3：正在播放）


#### 一、	菜单控件（Menu）：用于显示应用程序的自定义菜单。
###### 常用属性：
- Caption：菜单的标题文字
- Check：是否在菜单旁边显示复选标记。（布尔型）
- Enabled：是否响应用户操作。
- Index：索引（用于区分数组内的各个菜单控件）
- Name：名称
- Shortcut：快捷键
- visible:是否可见
- WindowList:是否维护当前MDI子窗口的列表。
###### 常用事件：
- 菜单控件只有一个Click事件。
##### 二、RichTextBox控件：富文本框
1、不仅允许输入和编辑文本，同时还提供了标准文本框控件所没有的、更高级的指定格式的许多功能。（√）
2、常用属性： 
- FileName：文件名
- MaxLength：最大字符数
- MultiLine：是否允许多行
- RightMargin：文本右边距
- ScrollBars：滚动条（取值 0:无,  1:水,  2:垂,  3:水和垂）
- SelAlignment：段落的对齐方式（取值：0：左  1：右   2：中）
- SelBold、SelItalic、SelStrikethru、SelUnderLine：
粗体、  斜体、    删除线、     下划线
- SelBullet：是否有项目符号
- SelCharoffset：决定文本是①在基线上 ②在基线之上（作上标）
- 在基线之下（作下标）
- SelColor:所选文本颜色
- SelFontName：所选文本的字体
- SelFontSize：所选文本的字号
- SelHangingIndent、SelIndent、SelRightIndent:所选文本的
悬挂缩进、       缩进及      右缩进状态
- SelRTF：当前选择的RTF文本
- SelTabCount、SelTabs:所选文本的制表符数目及制表符位置
- TextRTF:所有RTF文本
3、常用方法：
- Find：搜索文本
- GetLineFromChar：获取行号
- LoadFile：加载文件【2011年高考-单选题】
- SaveFile：保存文件【2015年高考-单选题】
- SelPrint：打印文件
#### 三、状态栏控件（StatusBar）:提供窗体，该窗体通常位于父窗体的底部。（√）
- StatusBar对象最多能被分成16个Panel对象，每个Panel对象可包含文本和图片）（√）
#### 四、ClipBoard对象（剪贴板）：
即剪贴板，提供对系统剪贴板的访问，用于操作剪贴板上的文本和图形。

###### 常用方法：
- Clear：清除剪贴板内容
- GetData：返回图形
- GetText：返回文本
- SetData：设置图形（即将指定图片放到剪贴板上）
- SetText：设置文本（即将指定文本放到剪贴板上）
#### 五、工具栏控件（ToolBar）
可以在窗体上创建工具栏；为用户访问应用程序的最常用功能和命令提供了图形接口。
- 要制作工具栏，需要用到两个ActiveX控件：Toolbar和Imagelist。ToolBar提供所需的按钮，ImageList则为每个按钮提供图像。
- 包含一个Button对象集合，该对象被用来创建与应用程序相关联的工具栏。（常和ImageList控件一起使用）
######  常用属性：
- Buttons：Button对象的集合。
- ImageList：返回或设置与工具栏相关联的ImageList控件。
#### 六、图像列表控件（ImageList）：包含ListImage对象的集合，该集合中的每个对象都可以通过其索引或关键字被引用。
>ImageList控件不能独立使用
只是作为一个便于向其他控件提供图像的资料中心，常和ToolBar控件一起
使用。


#项目七 
#### 一、驱动器列表框:即 DriveListBox
- 用来显示用户系统中所有有效磁盘驱动器的
列表。
常用属性:
- Drive:驱动器 - List:驱动器列表
- ListCount:驱动器个数 - ListIndex:驱动器索引
 常用事件:
- Change:当改变所选的驱动器时发生。
#### 二、目录列表框：即 DirListBox
在运行时显示目录和路径,该控件可用于显示分层的目录列表。
#####  常用属性:
- List:目录列表 - ListCount:子目录个数
- ListIndex:路径的索引 
- Path:当前路径
 ##### 常用事件:
- Change:当双击一个新的目录来改变所选目录或通过代码改变 Path 属性的
值时发生。
#### 三、文件列表框：即 FileListBox
在运行时把 Path 属性指定的目录中的文件显示出来，该控件用来显示所选择文件类型的文件列表。
 ###### 常用属性:
- FileName：路径和文件名(即包括路径的文件名)
- List：文件列表
- ListCount：文件个数
- ListIndex：文件的索引
- MultiSelect：能否多选(0:不能复选 1:简单复选 2:扩展复选)
- Path ：当前路径
- Pattern：指定文件格式(即文件类型，如*.frm,*.bmp 等)
 ###### 常用事件:
- Click：单击一个文件时发生
- PathChange：路径改变时发生
- PatternChange：样式改变时发生
 文件的访问方式 及 顺序文件的访问
#### 一、文件的访问方式：
- 顺序型,适于读写在连续块中的文本文件；
- 随机型，适于读写有固定长度记录结构的文本文件或二进制文件；
- 二进制型，适用于读写任意有结构的文件。
> 当要处理只包含文本的文件时，使用顺序型访问最好。

#### 二、文件的打开及读写操作（打开文件均用 Open 语句）
 顺序文件的打开方式
- Input 是只读方式
- Output 是只写方式
- Append 是追加记录的方式。
 打开二进制文件的 Open 语句语法:
`Open Pathname [For Binary] As filenumbr`
- 文件号的范围 1~511，缓冲字符数 buffersize <= 32767
- 当以 Input 方式打开顺序文件时，该文件必须已经存在,否则会出错。
- 当以Append、OutPut方式打开顺序文件时,如果Pathname指定的文件不存在, 则 Open 语句会先创建该文件,然后再打开它。
- 在以 Input、Append 或 OutPut 方式打开一个文件后，为其他类型的操作重新打
开它之前必须先用 Close 语句关闭它。 
- Close 语句的语法格式为:
`Close [[#]Filenumber] [，[#]Filenumber]…`

例如: close #1,#2,#3 （不能用#1~#3 或 #1-#3 表示）
- 如果省略文件号,将关闭打开的所有活动文件。即： Close (将关闭所有活动文件)
#### 三、顺序文件的读操作
 要检索文本文件的内容，应以 Input 方式打开
- LOF 函数:文件大小  （读作：Length of）
- EOF 函数:是否到达文件结尾（读作：End of）
######  访问与管理文件
顺序文件的写操作
- 要在顺序文件中存储变量的内容，应先以 Output 或 Append 方式打开它。
- 用 Print#语句或 Write#语句进行写操作
 顺序文件的关闭：
- 用 Close 语句(关闭文件均用 Close 语句)

###### 随机文件读写操作步骤（定 打 读写 关）
 1. 定义记录类型和变量；
 2. 使用 Open 语句以随机方式打开文件；
 3. 对记录进行读写操作； 
 4. 关闭随机文件。
#### 二、定义记录类型和变量
-  在 VisualBasic 中，用户可以自定义数据类型。
 - 在模块级别中用 Type 语句定义记录类型。
 Type 语句的语法格式为： 
`Type 记录名
…
End Type`
#### 三、随机文件的打开：
` 语法格式为：Open Pathname [For Random] As filenumbr`
> Random 是默认的访问类型,所以在打开随机文件时,For Random 可以省略。

#### 四、随机文件的读操作：
 用 Get 语句，语法格式为： （记忆方法：随机概读）
`Get [#]filenumber,[recnumber],varname`
	文件号 记录号 变量名 （参数顺序：文、记、变）
#### 五、随机文件的写操作: （P202）
 用 Put 语句，语法格式为: （记忆方法：随机谱写）
`Put [#]filenumber,[recnumber],varname`
	文件号 记录号 变量名 （参数顺序：文、记、变）
六、清除随机文件中删除的记录的步骤？（创 复 关删 命）
 1. 创建一个新文件；
2. 把有用的所有记录从原文件复制到新文件；
 3.关闭原文件并用 Kill 语句删除它；
 4.使用 Name 语句把新文件以原文件的名字重新命名。
知识补充：
**Name 语句**的语法格式为：
`Name 原文件 as 新文件`
#### FSO 对象模型

-  FSO 对象模型：提供了一个基于对象的工具来处理文件夹和文件。
- FSO 对象模型包含了哪些对象？各有何功能？
- Drive 对象：用于收集关于系统所用的驱动器的信息。
- Folder 对象:用于创建、移动或删除文件夹。
- Files 对象：用于创建、移动或删除文件。
- FileSystemObject 对象：用于创建、删除和收集相关信息，以及操作驱动器、文件夹和文件。
- TextStream 对象：用于读写文本文件。


 
#### 项目八 重点知识汇总

##### Data 控件：即数据控件，它是数据库与 Visual Basic 窗体之间的桥梁。
- Connect ：连接
- RecordSouse ：记录源
- DataBaseName ：数据库文件名
- RecordSetType ：记录集类型（取值：0 表 1 动 2 快照）
- ReadOnly：只读
- EOFAction ：结尾操作
- BOFAction ：开头操作
- Exclusive ：是否独占数据库（默认值 false）
- Select 语句：指定输出字段
- Form 子句 ：指定数据来源
- Where 子句：指定过滤条件
- Order By 子句 ：排序
- Group By 子句 ：分组
- Insert 语句：添加一行新记录。
- ODBC ：开放数据库互联
- API ： 应用程序编程接口
#### MSFlexGrid 控件
##### 常用属性
- AllowBigSelection：是否可使整行 - 整列都被选中
- AllowUserResizing：是否可对行和列的大小进行重新调整
- BackColorBand：带区的背景色
- BackColorHeader：标头区域的背景色
- BackColorindent： 缩进区域的背景色
- BackColorUnpopulated：未填充数据区域的背景色
- CellAlignment： 单元格的对齐方式
- DataSource：数据源
- Col 和 Row：坐标 Col：列 （横向）Row：行
- ColPosition：列的位置
- RowPosition：行的位置
- cols：总列数
- Rows：总行数
- ColSel：起始列或终止列
- RowSel：起始行或终止列
- ColWidth：列宽
- FixedCols：固定列的总数
- FixedRows： 固定行的总数
- TextMatrix：任意一个单元的文本内容
#### 重点知识汇总

##### 常用方法：
- AddItem： 添加一行
- Clear：清除所有内容
- RemoveItem：删除一行
##### RecordSet 对象
常用属性：
- EOF（读作 End of）：指针指向最后一条记录之后时，值为True
- BOF（读作 Before of）：指针指向首条记录之前时，值为True
- RecordCount：记录个数
-  AbsolutePosition：当前记录的记录号（即绝对位置）【重难点】
- NoMatch：没有符合条件的时，值为 True。（不匹配）
- Fields：所有字段
常用方法：
- MoveFirst：指针定位到第一条记录
- MoveLast：指针定位到最后一条记录
- MoveNext：指针定位到下一条记录
- MovePrevious：指针定位到上一条记录
- Move[n]：向前或向后移 n 条记录
- FindFirst：从开头查找满足条件的第一条记录
- IndLast：从末尾查找满足条件的第一条记录
- FindNext：从当前记录开始查找满足条件的下一条记录
- FindPrevious：从当前记录开始查找满足条件的上一条记录
- Update：更新记录
- AddNew：添加新记录
- Delete：删除当前记录
-   Edit：编辑
####  ADOConnection
- CursorLocation：游标位置
- DefaultDataBase：默认数据库
- Close：关闭
- Open：打开
- CursorType：游标类型
取值： 
adopenForwardOnly：仅向前游标
adopenKeyset：键集游标
adopenDynamic：动态游标
adopenStatic：静态游标
 RecordCount：记录数目
