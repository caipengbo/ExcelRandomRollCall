
// excelDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "CApplication.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"

#include "excel.h"
#include "excelDlg.h"
#include "afxdialogex.h"
#include <stdlib.h>
#include <time.h>
#include <windows.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CexcelDlg 对话框



CexcelDlg::CexcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_EXCEL_DIALOG, pParent)
	, m_selectCol(0)
	, m_showCellText(_T(""))
	, m_startRow(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_itemNum = -1;
}

void CexcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_selectCol);
	DDX_Text(pDX, IDC_EDIT2, m_showCellText);
	DDX_Text(pDX, IDC_EDIT3, m_startRow);
	DDX_Control(pDX, IDC_EDIT2, m_controlEdit);
}

BEGIN_MESSAGE_MAP(CexcelDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CexcelDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CexcelDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CexcelDlg 消息处理程序

BOOL CexcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	//随机数生成种子
	srand((unsigned int)time(0));
	m_Font.CreatePointFont(350, _T("宋体"), NULL);
	m_controlEdit.SetFont(&m_Font, FALSE);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CexcelDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CexcelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CexcelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


//导入excel文件
void CexcelDlg::OnBnClickedButton1()
{
	
	//UpdateData(TRUE) == 将控件的值赋值给成员变量;UpdateData(FALSE) == 将成员变量的值赋值给控件
	UpdateData(TRUE);
	CApplication app;
	CRange range;
	CWorkbook book;
	CWorkbooks books;
	CWorksheet sheet;
	CWorksheets sheets;
	LPDISPATCH lpdisp;
	COleVariant vresult;
	COleVariant covtrue((short)TRUE);
	COleVariant covfalse((short)FALSE);
	COleVariant covoptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	//创建服务
	if (!app.CreateDispatch(L"Excel.Application"))
	{
		AfxMessageBox(L"不能启动Excel");
		return;
	}
	app.put_Visible(FALSE);
	books.AttachDispatch(app.get_Workbooks());

	//打开文件
	CString FilePathName;
	CFileDialog dlg(TRUE);///TRUE为OPEN对话框，FALSE为SAVE AS对话框
	if (dlg.DoModal() == IDOK)
	{
		FilePathName = dlg.GetPathName();
		/*
		(1)GetPathName();取文件名全称，包括完整路径。取回C:\\WINDOWS\\TEST.EXE
		(2)GetFileTitle();取回TEST
		(3)GetFileName();取文件全名：TEST.EXE
		(4)GetFileExt();取扩展名EXE
		*/
	}
	else
	{
		return;
	}
	lpdisp = books.Open(
		FilePathName,
		covoptional, covfalse, covoptional, covoptional, covoptional,
		covoptional, covoptional, covoptional, covoptional, covoptional,
		covoptional, covoptional, covoptional, covoptional);

	//获得
	book.AttachDispatch(lpdisp);
	sheets.AttachDispatch(book.get_Worksheets());

	lpdisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpdisp);

	//获得区域 
	CRange usedrange;
	usedrange.AttachDispatch(sheet.get_UsedRange());

	//行数目 
	range.AttachDispatch(usedrange.get_Rows());
	long irownum = range.get_Count();
	long istartrow = usedrange.get_Row();
	

	//列数目
	range.AttachDispatch(usedrange.get_Columns());
	long icolnum = range.get_Count();
	long istartcol = usedrange.get_Column();

	m_itemNum = 0;
	if (m_startRow>0&&m_startRow<=(int)irownum&&m_selectCol>0&&m_selectCol<=(int)icolnum)
	{
		for (int i = m_startRow; i <= irownum; i++)
		{
			//获得单元格内容    get_Item(COleVariant i , COleVariant j)   i 行  j 列
			//特定列j
			range.AttachDispatch(usedrange.get_Item(  COleVariant((long)i), COleVariant(m_selectCol) ).pdispVal);
			VARIANT varItemName = range.get_Text();
			CString strItem = varItemName.bstrVal;
			//将数据存入vector之中 
			m_allItem.insert(m_allItem.begin()+ m_itemNum,strItem);
			++m_itemNum;
		}
		AfxMessageBox(L"导入成功");
	}
	else
	{
		AfxMessageBox(L"请输入有效的起始行与随机列");
	}
	
	/*
	for (int j = 0; j < m_itemNum; j++)
	{
		AfxMessageBox(m_allItem.at(j));
	}
	*/
	book.Close(covoptional, COleVariant(FilePathName), covoptional);
	books.Close();
	range.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	app.Quit();

}

//产生随机数
void CexcelDlg::OnBnClickedButton2()
{
	//生成10个备选项，可以存在相同的项
	UpdateData(TRUE);
	if (m_itemNum>=0)
	{
		int ready[10] = { 0 };
		for (int i = 0; i < 10; i++)
		{
			ready[i] = rand() % m_itemNum;
		}
		//随机显示框缓慢移动项目5秒
		for (int i = 0; i < 10; i++)
		{
			m_showCellText = m_allItem.at(ready[i]);
			UpdateData(FALSE);
			UpdateWindow();
			Sleep(400);
		}
		int finallyItem = rand() % 10;
		m_showCellText = m_allItem.at(finallyItem);
		UpdateData(FALSE);
		UpdateWindow();
		//最后在10个备选项里面 ，再随机产生一个最终显示项
		MessageBox(m_showCellText,L"恭喜你");
	}
	else
	{
		AfxMessageBox(L"请导入Excel文件");
	}
	
}
