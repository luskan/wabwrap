#pragma once

class CLVEdit : public CEdit
{
// Construction
public:
	CLVEdit() :m_x(0),m_y(0) {}

// Attributes
public:
	int m_x;
	int m_y;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CLVEdit)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CLVEdit() {};

	// Generated message _STL::map functions
protected:
	//{{AFX_MSG(CLVEdit)
	afx_msg void OnWindowPosChanging(WINDOWPOS FAR* lpwndpos);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

extern const UINT NEAR WM_LCTRL_LABELENDEDIT;

// CCustomListCtrl

class CCustomListCtrl : public CListCtrl
{
public:
  CCustomListCtrl();
protected: // create from serialization only

	DECLARE_DYNCREATE(CCustomListCtrl)

// Attributes
protected:
	int		m_cx;
	BOOL		m_nEdit;
	CLVEdit	m_LVEdit;

// Operations
protected:
	virtual void RepaintSelectedItems();
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDIS);
	virtual LPCTSTR MakeShortString(CDC* pDC,LPCTSTR lpszLong,int nColumnLen);

public:

	public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);

// Implementation
public:
	virtual ~CCustomListCtrl();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message _STL::map functions
protected:

  // Helpers
  void BeginLabelEdit( INT iItem );

	afx_msg void OnPaint();
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSize(UINT nType,int cx,int cy);
  afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnEndLabelEdit(NMHDR* pNMHDR,LRESULT* pResult);
	afx_msg void OnBeginLabelEdit(NMHDR* pNMHDR,LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
};


