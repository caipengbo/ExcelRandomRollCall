
// excelDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CexcelDlg �Ի���



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


// CexcelDlg ��Ϣ�������

BOOL CexcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	//�������������
	srand((unsigned int)time(0));
	m_Font.CreatePointFont(350, _T("����"), NULL);
	m_controlEdit.SetFont(&m_Font, FALSE);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CexcelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CexcelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


//����excel�ļ�
void CexcelDlg::OnBnClickedButton1()
{
	
	//UpdateData(TRUE) == ���ؼ���ֵ��ֵ����Ա����;UpdateData(FALSE) == ����Ա������ֵ��ֵ���ؼ�
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

	//��������
	if (!app.CreateDispatch(L"Excel.Application"))
	{
		AfxMessageBox(L"��������Excel");
		return;
	}
	app.put_Visible(FALSE);
	books.AttachDispatch(app.get_Workbooks());

	//���ļ�
	CString FilePathName;
	CFileDialog dlg(TRUE);///TRUEΪOPEN�Ի���FALSEΪSAVE AS�Ի���
	if (dlg.DoModal() == IDOK)
	{
		FilePathName = dlg.GetPathName();
		/*
		(1)GetPathName();ȡ�ļ���ȫ�ƣ���������·����ȡ��C:\\WINDOWS\\TEST.EXE
		(2)GetFileTitle();ȡ��TEST
		(3)GetFileName();ȡ�ļ�ȫ����TEST.EXE
		(4)GetFileExt();ȡ��չ��EXE
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

	//���
	book.AttachDispatch(lpdisp);
	sheets.AttachDispatch(book.get_Worksheets());

	lpdisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpdisp);

	//������� 
	CRange usedrange;
	usedrange.AttachDispatch(sheet.get_UsedRange());

	//����Ŀ 
	range.AttachDispatch(usedrange.get_Rows());
	long irownum = range.get_Count();
	long istartrow = usedrange.get_Row();
	

	//����Ŀ
	range.AttachDispatch(usedrange.get_Columns());
	long icolnum = range.get_Count();
	long istartcol = usedrange.get_Column();

	m_itemNum = 0;
	if (m_startRow>0&&m_startRow<=(int)irownum&&m_selectCol>0&&m_selectCol<=(int)icolnum)
	{
		for (int i = m_startRow; i <= irownum; i++)
		{
			//��õ�Ԫ������    get_Item(COleVariant i , COleVariant j)   i ��  j ��
			//�ض���j
			range.AttachDispatch(usedrange.get_Item(  COleVariant((long)i), COleVariant(m_selectCol) ).pdispVal);
			VARIANT varItemName = range.get_Text();
			CString strItem = varItemName.bstrVal;
			//�����ݴ���vector֮�� 
			m_allItem.insert(m_allItem.begin()+ m_itemNum,strItem);
			++m_itemNum;
		}
		AfxMessageBox(L"����ɹ�");
	}
	else
	{
		AfxMessageBox(L"��������Ч����ʼ���������");
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

//���������
void CexcelDlg::OnBnClickedButton2()
{
	//����10����ѡ����Դ�����ͬ����
	UpdateData(TRUE);
	if (m_itemNum>=0)
	{
		int ready[10] = { 0 };
		for (int i = 0; i < 10; i++)
		{
			ready[i] = rand() % m_itemNum;
		}
		//�����ʾ�����ƶ���Ŀ5��
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
		//�����10����ѡ������ �����������һ��������ʾ��
		MessageBox(m_showCellText,L"��ϲ��");
	}
	else
	{
		AfxMessageBox(L"�뵼��Excel�ļ�");
	}
	
}
