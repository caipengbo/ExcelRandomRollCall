
// excelDlg.h : 头文件
//
#include <vector>
#include "afxwin.h"
using namespace std;
#pragma once


// CexcelDlg 对话框
class CexcelDlg : public CDialogEx
{
// 构造
public:
	CexcelDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCEL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	long m_selectCol;
	CString m_showCellText;
	afx_msg void OnBnClickedButton2();
	int m_itemNum;
	vector<CString> m_allItem;
	long m_startRow;
	CFont m_Font;
	CEdit m_controlEdit;
};
