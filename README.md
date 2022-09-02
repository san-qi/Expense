# Expense

基于 C#的 WPF 框架所编写的 Excel 文件生成器,是继大四毕设之后的又一玩具巨作  
本来是准备用 Rust 来写的 ~~我老 Rust 了~~,但考虑到这是 Windows 系统所用,所以就......  
欸,结果只能说差强人意  
后续再完善叭,谁知道有后续没呢

## 项目结构

基本是随心所欲,没刻意考虑设计模式  
项目重构了很多次,最根本的原因是生成一个 Excel 文件很繁琐,而且 NPOI 框架封装的并不算很完美,大量代码需要进一步抽象

```
.
├── Expense
│   ├── App.xaml
│   ├── App.xaml.cs
│   ├── AssemblyInfo.cs
│   ├── Book.cs
│   ├── Expense.csproj
│   ├── Expense.sln
│   ├── main.ico
│   ├── MainWindow.xaml
│   ├── MainWindow.xaml.cs
│   ├── Program.cs
│   ├── RangeCell.cs
│   ├── Reception.cs
│   ├── Trip.cs
│   └── 集萤映雪.ttf
├── LICENSE.txt
└── README.md

1 directory, 16 files
```

- App.\*、MainWindow.\*:
  - xaml 文件涉及 WPF(即 GUI)部分的前端表现及相应的响应操作
    - App.xaml 中主要包含 Comboox 组件的一个重写,因为巨硬自带的这一组件样式必须通过重写来変更外观
    - Expense.xaml 中包含整个页面的展示逻辑,文件上半部分的样式可另开一个文件,但是懒癌犯了
    - 不得不说 WPF 的理念在现在来看也还是很优秀的,更别提这是一个很早之前的框架了;用它写出来的界面效果几乎完美,可定制化极高,当然代价你懂得
  - App.cs 与 MainWindow.cs 为相应前端的响应逻辑
- RangeCell.cs: 用于辅助表示单元格的类
- Program.cs: Excel 文件生成的调用入口,并没有什么参考价值
- Book.cs: 抽象出来的 Excel 文件生成器基类,包含很多生成时所需的自编写函数以及一些抽象共用逻辑
- Reception.cs: 具体的 Reception 相关 Excel 文件生成子类
- Trip.cs: 具体的 Trip 相关 Excel 文件生成子类
- 其他资源文件

## 实际展示

TODO

GUI 部分是几天内做出来的,但效果不错  
唯一可惜的是生成的 EXE 文件可移植性不强,还需要 Dotnet 环境/恼,评价是不如 Rust

## 授权许可

既然是玩具,当然可以随便拿去玩啦,但是具体请遵循[GNU GPLv3](https://www.gnu.org/licenses/gpl-3.0.html)许可
