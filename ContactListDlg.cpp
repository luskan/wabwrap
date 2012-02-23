// ContactListDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ContactList.h"
#include "ContactListDlg.h"
#include ".\contactlistdlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CContactListDlg dialog

const CRect CContactListDlg::rtInvalid = CRect(-1, 0, 0, 0);

CContactListDlg::CContactListDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CContactListDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
  rtClientOrg = rtInvalid;
  rtListOrg = rtInvalid;
}

void CContactListDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);
  DDX_Control(pDX, IDC_LIST1, ContactList);
  DDX_Control(pDX, IDC_FOLDERS, Folders);
  DDX_Control(pDX, IDC_EDIT1, edFilter);
}

BEGIN_MESSAGE_MAP(CContactListDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
  ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
  ON_CBN_SELCHANGE(IDC_FOLDERS, OnCbnSelchangeFolders)
  ON_EN_CHANGE(IDC_EDIT1, OnEnChangeEdit1)
  ON_REGISTERED_MESSAGE(WM_LCTRL_LABELENDEDIT, OnLabelEndEdit)
  ON_WM_SIZE()
END_MESSAGE_MAP()


// CContactListDlg message handlers

BOOL CContactListDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

  SetWindowText(_T("Wab Wraper"));  

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
  //
  // Specify fields to be shown in columns.
  WabWrap.Properties().Add(PR_OBJECT_TYPE);
  WabWrap.Properties().Add(PR_ENTRYID);
  WabWrap.Properties().Add(PR_DISPLAY_NAME);
  WabWrap.Properties().Add(PR_COMPANY_NAME);
  WabWrap.Properties().Add(PR_DEPARTMENT_NAME);

  WabWrap.Properties().Add(PR_GIVEN_NAME);
  WabWrap.Properties().Add(PR_MIDDLE_NAME);
  WabWrap.Properties().Add(PR_SURNAME);
  WabWrap.Properties().Add(PR_STREET_ADDRESS);

  WabWrap.Properties().Add(PR_CONTACT_EMAIL_ADDRESSES);
  WabWrap.Properties().Add(PR_EMAIL_ADDRESS);

  WabWrap.Properties().Add(PR_STATE_OR_PROVINCE);
  WabWrap.Properties().Add(PR_POSTAL_CODE);
  WabWrap.Properties().Add(PR_COUNTRY);

  WabWrap.Properties().Add(PR_HOME_ADDRESS_STREET);
  WabWrap.Properties().Add(PR_HOME_ADDRESS_CITY);
  WabWrap.Properties().Add(PR_HOME_ADDRESS_STATE_OR_PROVINCE);
  WabWrap.Properties().Add(PR_HOME_ADDRESS_POSTAL_CODE);
  WabWrap.Properties().Add(PR_HOME_ADDRESS_COUNTRY);

  //
  // Connect to Windows Address Book and set initially PAB folder.
  WabWrap.Connect(this->GetSafeHwnd());     
  
  // Setup selection in combobox with available folders.
  Folders.SetCurSel(0);

  // Setup columns of contact list.
  ContactList.SetExtendedStyle(LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES);
  ContactList.InsertColumn(0, _T("#"), LVCFMT_LEFT, 20);
  ContactList.InsertColumn(1, _T("Type"), LVCFMT_LEFT, 50);
  ContactList.InsertColumn(2, _T("Contact name"), LVCFMT_LEFT, 120);
  ContactList.InsertColumn(3, _T("Emails"), LVCFMT_LEFT, 120);
  ContactList.InsertColumn(4, _T("First name"), LVCFMT_LEFT, 80);
  ContactList.InsertColumn(5, _T("Last name"), LVCFMT_LEFT, 80);
  ContactList.InsertColumn(6, _T("City"), LVCFMT_LEFT, 80);
  ContactList.InsertColumn(7, _T("Street"), LVCFMT_LEFT, 80);
  ContactList.InsertColumn(8, _T("Country"), LVCFMT_LEFT, 80);  

  // Fill rows in list with contacts for currently selected folder.
  UpdateFieldData();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CContactListDlg::UpdateFieldData() {

  /// Update text control with number of elements in current folder.

  CWnd* pConCount = GetDlgItem(IDC_CONTACT_COUNT);
  CString sCount;
  INT iCount = WabWrap.GetFolderItemCount();
  if (iCount == -1)
    sCount = _T("Contact count : 0");
  else
    sCount.Format(_T("Contact count : %d"), iCount);
  pConCount->SetWindowText(sCount);

  //CContact con;
  //con.bDistList = true;
  //WabWrap.UpdateContact(con);

  /// Update list.

  ContactList.DeleteAllItems();
  int iItemType;
  for (int i=0; i<iCount; i++) {
        
    /// List item number column.
    CString sNum;
    sNum.Format(_T("%d"), i);
    ContactList.InsertItem(i, sNum );
    

    /// Get folder item type.
    iItemType = WabWrap.GetItemType(i);

    if (iItemType == CWABWrapper::DIST_LIST) {
      CDistList DistList;
      CMailUser MailUser;
      WabWrap.GetDistList(i, DistList);

      /*
      /// Iterate all distribution list items, and dump them to OUTPUT window.
      for (int j=0; j<DistList.aMembers.size(); ++j) {        
        if (WabWrap.GetItemType(DistList.aMembers[j]) == CWABWrapper::MAIL_USER) {
          WabWrap.GetMailUser(DistList.aMembers[j], MailUser);
          TRACE(_T("Member %s\n"), MailUser.sFirstName.c_str());
        } 
        else {
          TRACE(_T("Its a distribution list\n"));
        }
      }
      */
      ContactList.SetItemText(i, 1, _T("DistList"));
      ContactList.SetItemText(i, 2, DistList.GetStringPropValue(PR_DISPLAY_NAME).c_str());
      ContactList.SetItemText(i, 3, _T("--"));
      ContactList.SetItemText(i, 4, _T("--"));
      ContactList.SetItemText(i, 5, _T("--"));
      ContactList.SetItemText(i, 6, DistList.GetStringPropValue(PR_HOME_ADDRESS_CITY).c_str());
      ContactList.SetItemText(i, 7, DistList.GetStringPropValue(PR_HOME_ADDRESS_STREET).c_str());
      ContactList.SetItemText(i, 8, DistList.GetStringPropValue(PR_HOME_ADDRESS_COUNTRY).c_str());
    } 
    else
    if (CWABWrapper::MAIL_USER == iItemType) {      
      CMailUser MailUser;
      WabWrap.GetMailUser(i, MailUser);
      ContactList.SetItemText(i, 1, _T("Mail"));
      ContactList.SetItemText(i, 2, MailUser.GetStringPropValue(PR_DISPLAY_NAME).c_str());
      
      _STL::string sAllEmails;
      _STL::vector<_STL::string> aEmails;
      HRESULT hr = MailUser.GetStringPropMultiValue(PR_CONTACT_EMAIL_ADDRESSES, aEmails);
      ASSERT(hr == S_OK);
      for (int n=0; n<aEmails.size(); n++){
        sAllEmails+=aEmails[n];
        if (n+1<aEmails.size())
          sAllEmails+=_T(",");
      }
      ContactList.SetItemText(i, 3, sAllEmails.c_str());

      ContactList.SetItemText(i, 4, MailUser.GetStringPropValue(PR_GIVEN_NAME).c_str());
      ContactList.SetItemText(i, 5, MailUser.GetStringPropValue(PR_SURNAME).c_str());
      ContactList.SetItemText(i, 6, MailUser.GetStringPropValue(PR_HOME_ADDRESS_CITY).c_str());
      ContactList.SetItemText(i, 7, MailUser.GetStringPropValue(PR_HOME_ADDRESS_STREET).c_str());
      ContactList.SetItemText(i, 8, MailUser.GetStringPropValue(PR_HOME_ADDRESS_COUNTRY).c_str());

    } 
    else {
      ASSERT(0);
    } 
  }

  /// Last item is for adding new contact/ditlist.
  CString sNum;
  sNum.Format(_T("*"));
  ContactList.InsertItem(iCount, sNum );
  ContactList.SetItemText(iCount, 1, _T("Click here to add new contact/distlist"));

  /// Fill combobox with known WAB folders.
  _STL::vector<_STL::string> aFolders;
  INT iCurSel = Folders.GetCurSel();
  Folders.ResetContent();
  Folders.AddString(_T("All"));
  if (SUCCEEDED(WabWrap.GetFolders(aFolders))) {
    for(size_t i=0; i<aFolders.size(); i++) 
      Folders.AddString(aFolders[i].c_str());
  }
  Folders.SetCurSel(iCurSel == -1 ? 0 : iCurSel);

//  WabWrap.CreateFolder("Test21");
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CContactListDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CContactListDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CContactListDlg::OnBnClickedCancel()
{
  // TODO: Add your control notification handler code here
  OnCancel();
}

void CContactListDlg::OnCbnSelchangeFolders()
{
  int iCurSel;
  if((iCurSel = Folders.GetCurSel()) != -1) {
    CString sFolder;
    _STL::string sStlStr;

    if (iCurSel==0) {
      WabWrap.SetPABFolder();
    } 
    else {
      Folders.GetLBText(iCurSel, sFolder);    
      sFolder.ReleaseBuffer();
      sStlStr = sFolder.GetBuffer(255);
      WabWrap.SetFolder(sStlStr);
    }

    UpdateFieldData();
    Folders.SetCurSel(iCurSel);
  }
}

void CContactListDlg::UpdateListWithFiltering() {
  
  // Get filter _STL::string
  CString sFilter;
  edFilter.GetWindowText(sFilter);
  
  // Filter currently selected folder.
  _STL::string sStlStr = sFilter.GetBuffer(255);
  
  //CDynamicPropertyTagArray FilterProps;
  //FilterProps.Add(PR_DISPLAY_NAME);
  //HRESULT hr = WabWrap.SetFilter(sStlStr, FilterProps);

  HRESULT hr = WabWrap.SetFilter(sStlStr);
  if (FAILED(hr))
    SetWindowText(_T("Filter Error"));
  else
    SetWindowText(_T("WabWraper"));    
  sFilter.ReleaseBuffer();

  //
  UpdateFieldData();
}

void CContactListDlg::OnEnChangeEdit1() {
  UpdateListWithFiltering();
}

LRESULT CContactListDlg::OnLabelEndEdit(WPARAM wParam, LPARAM lParam)
{
  UNREFERENCED_PARAMETER(lParam); 
  //UpdateMailUser

  INT iItem = LOWORD(wParam);
  INT iSubItem = HIWORD(wParam);

  _STL::string sNewText = (LPCTSTR)ContactList.GetItemText(iItem, iSubItem);
  INT iItemCount = WabWrap.GetFolderItemCount();

  if ( iItem == iItemCount ) {
    /// Its New Contact item.    
    CMailUser MailUser;
    
    switch ( iSubItem ) {
      case 2:
        MailUser.AddStringPropValue(PR_DISPLAY_NAME, sNewText);
        break;
      case 4:
        MailUser.AddStringPropValue(PR_GIVEN_NAME, sNewText);
      break;        
    };

    HRESULT hr = WabWrap.AddMailUser(MailUser);    
    if ( FAILED(hr) )
      AfxMessageBox(WabWrap.GetMAPIError(hr).c_str());
    UpdateListWithFiltering();
  }
  else {
    INT iItemType = WabWrap.GetItemType(iItem);
    if ( CWABWrapper::MAIL_USER == iItemType ) {
      CMailUser MailUser;
      WabWrap.GetMailUser(iItem, MailUser);
      
      switch ( iSubItem ) {
      case 2:
        MailUser.AddStringPropValue(PR_DISPLAY_NAME, sNewText);
        break;        
      case 4:
        MailUser.AddStringPropValue(PR_GIVEN_NAME, sNewText);
        break;        
      };
      WabWrap.UpdateMailUser(MailUser);
    }
  }

  //ContactList.SetItemText(i, 1, _T("DistList"));
  
  /*ContactList.SetItemText(i, 3, _T("--"));
  ContactList.SetItemText(i, 4, _T("--"));
  ContactList.SetItemText(i, 5, _T("--"));
  ContactList.SetItemText(i, 6, DistList.sCity.c_str());
  ContactList.SetItemText(i, 7, DistList.sStreet.c_str());
  ContactList.SetItemText(i, 8, DistList.sCountry.c_str());*/

  TRACE(_T("Editing label item: %d, subitem: %d"), LOWORD(wParam), HIWORD(wParam));
  return 0;
}

BOOL CContactListDlg::PreTranslateMessage(MSG* pMsg)
{
  /*if (pMsg->message == WM_KEYDOWN && pMsg->wParam == 'D') {
    CDistList dl;
    dl.sName = _T("My test d");
    WabWrap.AddDistList(dl);
    MessageBox(_T("Test"));
    UpdateFieldData();
  }*/
  return CDialog::PreTranslateMessage(pMsg);
}

void CContactListDlg::OnSize(UINT nType, int cx, int cy)
{
  CDialog::OnSize(nType, cx, cy);
  
  CRect rtClient;
  GetClientRect(&rtClient);
  
  CRect rtList;
  ContactList.GetWindowRect(&rtList);
  ScreenToClient(&rtList);

  if ( rtInvalid == rtClientOrg || rtInvalid == rtListOrg ) {
    rtClientOrg = rtClient;
    rtListOrg = rtList;
  }
  else {
    CRect rtNewList = rtListOrg;
    rtNewList.right = rtClient.right; 
    rtNewList.right -= rtClientOrg.right - rtListOrg.right;
    rtNewList.bottom = rtClient.bottom; 
    rtNewList.bottom -= rtClientOrg.bottom - rtListOrg.bottom;
    ContactList.MoveWindow(&rtNewList);
  }

  CRect rtStatic;
  GetDlgItem(IDC_CONTACT_COUNT)->GetWindowRect(&rtStatic);
  ScreenToClient(&rtStatic);
  rtStatic.top = rtClient.bottom - (rtClientOrg.bottom - rtListOrg.bottom);
  rtStatic.bottom = rtClient.bottom;
  GetDlgItem(IDC_CONTACT_COUNT)->MoveWindow(&rtStatic);
  
  GetDlgItem(IDC_INFO)->GetWindowRect(&rtStatic);
  ScreenToClient(&rtStatic);
  rtStatic.top = rtClient.bottom - (rtClientOrg.bottom - rtListOrg.bottom);
  rtStatic.bottom = rtClient.bottom;
  GetDlgItem(IDC_INFO)->MoveWindow(&rtStatic);  
}
