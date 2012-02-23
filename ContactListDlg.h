// ContactListDlg.h : header file
//

#pragma once

#include "WABWrapper.h"
#include "afxcmn.h"
#include "afxwin.h"
#include "CustomListCtrl.h"

// CContactListDlg dialog
class CContactListDlg : public CDialog
{
protected:
  void UpdateFieldData();

// Construction
public:
	CContactListDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_CONTACTLIST_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;
  CWABWrapper WabWrap;

  void UpdateListWithFiltering();

	// Generated message _STL::map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
  CCustomListCtrl ContactList;
  afx_msg void OnBnClickedCancel();
  CComboBox Folders;
  afx_msg void OnCbnSelchangeFolders();
  afx_msg void OnEnChangeEdit1();
  CEdit edFilter;
  afx_msg LRESULT OnLabelEndEdit(WPARAM wParam, LPARAM lParam);
  virtual BOOL PreTranslateMessage(MSG* pMsg);
  afx_msg void OnSize(UINT nType, int cx, int cy);
private:
  const static CRect rtInvalid;

  // Initial value of client rect.
  CRect rtClientOrg;
  
  // Initial value of list rect.
  CRect rtListOrg;
};
