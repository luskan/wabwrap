// CustomListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "ContactList.h"
#include "CustomListCtrl.h"

const UINT WM_LCTRL_LABELENDEDIT = RegisterWindowMessage( _T("WM_LCTRL_LABELENDEDIT") );

/////////////////////////////////////////////////////////////////////////////
// CLVEdit

BEGIN_MESSAGE_MAP(CLVEdit, CEdit)
	//{{AFX_MSG_MAP(CLVEdit)
	ON_WM_WINDOWPOSCHANGING()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CLVEdit::OnWindowPosChanging(WINDOWPOS FAR* lpwndpos) 
{
	lpwndpos->x=m_x;
	lpwndpos->y=m_y;
	CEdit::OnWindowPosChanging(lpwndpos);
}

// CCustomListCtrl

IMPLEMENT_DYNAMIC(CCustomListCtrl, CListCtrl)

BEGIN_MESSAGE_MAP(CCustomListCtrl, CListCtrl)
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_SETFOCUS()
	ON_WM_KILLFOCUS()
  ON_WM_LBUTTONDBLCLK()
	ON_NOTIFY_REFLECT(LVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(LVN_BEGINLABELEDIT, OnBeginLabelEdit)	
END_MESSAGE_MAP()

CCustomListCtrl::CCustomListCtrl()
{
	m_cx=0;
	m_nEdit=-1;
}

CCustomListCtrl::~CCustomListCtrl()
{
}

BOOL CCustomListCtrl::PreCreateWindow(CREATESTRUCT& cs)
{
	cs.style|=LVS_REPORT|
				 LVS_SINGLESEL|
				 LVS_SHOWSELALWAYS|
				 LVS_EDITLABELS|
				 LVS_OWNERDRAWFIXED;

	return CListCtrl::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CCustomListCtrl diagnostics

#ifdef _DEBUG
void CCustomListCtrl::AssertValid() const
{
	CListCtrl::AssertValid();
}

void CCustomListCtrl::Dump(CDumpContext& dc) const
{
	CListCtrl::Dump(dc);
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CCustomListCtrl methods

LPCTSTR CCustomListCtrl::MakeShortString(CDC* pDC, LPCTSTR lpszLong, int nColumnLen)
{
	static _TCHAR szLVStaticTemp[1024];

	static const _TCHAR szThreeDots[]=_T("...");

	int nStringLen=lstrlen(lpszLong);

	if(nStringLen==0 || pDC->GetTextExtent(lpszLong,nStringLen).cx<=nColumnLen)
		return(lpszLong);

	lstrcpy(szLVStaticTemp,lpszLong);
	int nAddLen=pDC->GetTextExtent(szThreeDots,sizeof(szThreeDots)).cx;

	for(int i=nStringLen-1; i>0; i--)
	{
		szLVStaticTemp[i]=0;
		if(pDC->GetTextExtent(szLVStaticTemp,i).cx+nAddLen<=nColumnLen)
			break;
	}

	lstrcat(szLVStaticTemp,szThreeDots);

	return(szLVStaticTemp);
}

#define OFFSET_FIRST	2
#define OFFSET_OTHER	6

void CCustomListCtrl::DrawItem(LPDRAWITEMSTRUCT lpDIS)
{
	int nItem=lpDIS->itemID;

	CDC* pDC=CDC::FromHandle(lpDIS->hDC);
		
	CRect rcAllLabels;
	this->GetItemRect(nItem,rcAllLabels,LVIR_BOUNDS);

	CRect rcItem;
	this->GetItemRect(nItem,rcItem,LVIR_LABEL);

	rcAllLabels.left=rcItem.left;
	if(rcAllLabels.right<m_cx)
		rcAllLabels.right=m_cx-2;

	LV_ITEM lvi;
	lvi.mask=LVIF_STATE;
	lvi.iItem=nItem;
	lvi.iSubItem=0;
	lvi.stateMask=0xFFFF;		
	this->GetItem(&lvi);

	BOOL bFocus=(GetFocus()==this);
	BOOL bSelected=(lvi.state&LVIS_SELECTED && 
						(GetStyle()&LVS_SHOWSELALWAYS));						 
	bSelected=bSelected||(lvi.state&LVIS_DROPHILITED);

	int nOldDCMode=pDC->SaveDC();

	if(bSelected && m_nEdit==-1)
	{	
		pDC->SetBkColor(::GetSysColor(COLOR_HIGHLIGHT));
		pDC->SetTextColor(::GetSysColor(COLOR_HIGHLIGHTTEXT));
		pDC->FillRect(rcAllLabels,&CBrush(::GetSysColor(COLOR_HIGHLIGHT)));
	}
	else
	{	
		pDC->SetBkColor(::GetSysColor(COLOR_WINDOW));
		pDC->SetTextColor(::GetSysColor(COLOR_WINDOWTEXT));
		pDC->FillRect(rcAllLabels,&CBrush(::GetSysColor(COLOR_WINDOW)));
	}
	
	UINT nJustify=DT_LEFT;
	UINT nFormat=DT_SINGLELINE|DT_NOPREFIX|DT_VCENTER;

	CString sText;

	if(m_nEdit!=0)
	{
		sText=MakeShortString(pDC,this->GetItemText(nItem,0),rcItem.Width());
		pDC->DrawText(sText,-1,&rcItem,nFormat|nJustify);
	}

	rcItem.left+=OFFSET_FIRST;
	rcItem.right-=OFFSET_FIRST;

	LV_COLUMN lvc;
	lvc.mask=LVCF_FMT|LVCF_WIDTH;

	for(int nCol=1; this->GetColumn(nCol,&lvc); nCol++)
	{
		switch(lvc.fmt&LVCFMT_JUSTIFYMASK)
		{
			case LVCFMT_RIGHT:
				nJustify=DT_RIGHT;
				break;
			case LVCFMT_CENTER:
				nJustify=DT_CENTER;
				break;
			default:
				nJustify=DT_LEFT;
				break;
		}

		rcItem.left=rcItem.right;
		rcItem.right+=lvc.cx;

		rcItem.left+=OFFSET_OTHER;
		rcItem.right-=OFFSET_OTHER;

		if(m_nEdit!=nCol)
		{	
			sText=MakeShortString(pDC,this->GetItemText(nItem,nCol),rcItem.Width());
			pDC->DrawText(sText,-1,&rcItem,nFormat|nJustify);
		}

		rcItem.right+=OFFSET_OTHER;
	}
	
	if(lvi.state&LVIS_FOCUSED && bFocus)
		pDC->DrawFocusRect(rcAllLabels);

	pDC->RestoreDC(nOldDCMode);	
}

void CCustomListCtrl::RepaintSelectedItems()
{
	CRect rcItem;
	CRect rcLabel;

	CListCtrl* pList=this;

	int nItem=pList->GetNextItem(-1,LVNI_FOCUSED);

	if(nItem!=-1)
	{
		pList->GetItemRect(nItem,rcItem,LVIR_BOUNDS);
		pList->GetItemRect(nItem,rcLabel,LVIR_LABEL);
		rcItem.left=rcLabel.left;

		InvalidateRect(rcItem,FALSE);
	}

	if(!(GetStyle()&LVS_SHOWSELALWAYS))
	{
		for(nItem=pList->GetNextItem(-1,LVNI_SELECTED);
			 nItem!=-1; 
			 nItem=pList->GetNextItem(nItem,LVNI_SELECTED))
		{
			pList->GetItemRect(nItem,rcItem,LVIR_BOUNDS);
			pList->GetItemRect(nItem,rcLabel,LVIR_LABEL);
			rcItem.left=rcLabel.left;

			InvalidateRect(rcItem,FALSE);
		}
	}

	UpdateWindow();
}

/////////////////////////////////////////////////////////////////////////////
// CCustomListCtrl message handlers

void CCustomListCtrl::OnSize(UINT nType, int cx, int cy) 
{
	m_cx=cx;
	CListCtrl::OnSize(nType, cx, cy);
}

void CCustomListCtrl::OnPaint() 
{
	CRect rcAllLabels;
	CListCtrl* pList=this;

	pList->GetItemRect(0,rcAllLabels,LVIR_BOUNDS);

	if(rcAllLabels.right<m_cx)
	{
		CPaintDC dc(this);

		CRect rcClip;
		dc.GetClipBox(rcClip);

		rcClip.left=min(rcAllLabels.right-1,rcClip.left);
		rcClip.right=m_cx;

		InvalidateRect(rcClip,FALSE);
	}

	CListCtrl::OnPaint();
}

void CCustomListCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	CListCtrl::OnSetFocus(pOldWnd);

	if(pOldWnd!=NULL && pOldWnd->GetParent()==this)
		return;
	
	RepaintSelectedItems();	
}

void CCustomListCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	CListCtrl::OnKillFocus(pNewWnd);

	if(pNewWnd!=NULL && pNewWnd->GetParent()==this)
		return;

	RepaintSelectedItems();		
}

void CCustomListCtrl::OnLButtonDblClk(UINT nFlags, CPoint point) {
  LVHITTESTINFO lvhti;
  lvhti.pt = point;
  this->SubItemHitTest(&lvhti);
  if ( lvhti.flags & LVHT_ONITEMLABEL ) {
    SendMessage(LVM_EDITLABEL, lvhti.iItem, 0);
  }
}

void CCustomListCtrl::OnEndLabelEdit(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LV_DISPINFO* pDispInfo=(LV_DISPINFO*)pNMHDR;
	
	CString sEdit=pDispInfo->item.pszText;

	if(!sEdit.IsEmpty())
	{
		this->SetItemText(pDispInfo->item.iItem,
										  m_nEdit,
										  sEdit);
	}

	VERIFY(m_LVEdit.UnsubclassWindow()!=NULL);

	this->SetItemState(pDispInfo->item.iItem,0,LVNI_FOCUSED|LVNI_SELECTED);
  GetParent()->PostMessage(WM_LCTRL_LABELENDEDIT, MAKEWPARAM(pDispInfo->item.iItem, m_nEdit));
  m_nEdit=-1;
  *pResult=0;
}

void CCustomListCtrl::OnBeginLabelEdit(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LV_DISPINFO* pDispInfo=(LV_DISPINFO*)pNMHDR;

  BeginLabelEdit(pDispInfo->item.iItem);

	*pResult=0;
}

void CCustomListCtrl::BeginLabelEdit( INT iItem ) 
{
	CPoint posMouse;
	GetCursorPos(&posMouse);
	ScreenToClient(&posMouse);

	LV_COLUMN lvc;
	lvc.mask=LVCF_WIDTH;

	CRect rcItem;
	this->GetItemRect(iItem, rcItem, LVIR_LABEL);

	if(rcItem.PtInRect(posMouse))
		m_nEdit=0;	

	INT nCol=1;
	while(m_nEdit==-1 && this->GetColumn(nCol, &lvc)) {	
		rcItem.left = rcItem.right;
		rcItem.right += lvc.cx;

		if(rcItem.PtInRect(posMouse))
			m_nEdit=nCol;	

		nCol++;	
	}

	if(m_nEdit == -1)
		return;
			
	HWND hWnd=(HWND)SendMessage(LVM_GETEDITCONTROL);
	ASSERT(hWnd!=NULL);
	VERIFY(m_LVEdit.SubclassWindow(hWnd));

	m_LVEdit.m_x=rcItem.left;
	m_LVEdit.m_y=rcItem.top-1;

	m_LVEdit.SetWindowText(this->GetItemText(iItem, m_nEdit));
}