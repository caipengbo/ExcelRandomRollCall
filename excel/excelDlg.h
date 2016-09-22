
// excelDlg.h : ͷ�ļ�
//
#include <vector>
#include "afxwin.h"
using namespace std;
#pragma once


// CexcelDlg �Ի���
class CexcelDlg : public CDialogEx
{
// ����
public:
	CexcelDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCEL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
