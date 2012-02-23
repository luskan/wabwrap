
#include <initguid.h>
#define INITGUID
#define USES_IID_IMAPITable 
#define USES_IID_IABContainer 
#define USES_IID_IDistList
#define USES_IID_IMailUser

#include <assert.h>
#include <mapiguid.h>
#include "WABWrapper.h"

#pragma comment(lib, "mapi32.lib")

#define SAFE_RELEASE(a) \
  if (a!=NULL) {        \
    a->Release();       \
    a=NULL;             \
  }

static SizedSPropTagArray( 3, sptDistEntryIdCols ) = { 3,
  PR_OBJECT_TYPE,
  PR_ENTRYID,
  PR_DISPLAY_NAME
};

const int CWABWrapper::DIST_LIST = 0;
const int CWABWrapper::MAIL_USER = 1;

// --------------------------------------------------------

CWABWrapper::CWABWrapper() {
  lpAdrBook = NULL;
  lpContainer = NULL;
  lpWABObject = NULL;
  lpMAPISession = NULL;
  lpTable = NULL;
  hInstLib = NULL;
  bIsConnected = false;
}

CWABWrapper::~CWABWrapper() {
  ClearFolderMap();
  if (IsConnected())
    Disconnect();
}

// --------------------------------------------------------

void CWABWrapper::ClearFolderMap() {
  for(FolderMap::iterator itr=WabFolders.begin(); itr!=WabFolders.end(); ++itr) 
    delete (*itr).second;
  WabFolders.clear();
}

HRESULT CWABWrapper::Remove(mapi_TEntryid& EntryID) {
  return S_OK;
}

// --------------------------------------------------------

typedef HRESULT (WINAPI *fWABOpen)(LPADRBOOK*,LPWABOBJECT*,LPWAB_PARAM,DWORD);
HRESULT CWABWrapper::Connect(HWND p_hwnd, bool bEnableProfiles) {
  DWORD Reserved2 = NULL;
  HRESULT hr;

  //
  // Disconnect first if already connected.
  if (IsConnected()) Disconnect();

  // Initialize MAPI Subsystem.
  hr = MAPIInitialize(NULL);
  MAPI_ASSERT_EX(hr);

  /*
  // Logon to MAPI.
  hr = MAPILogonEx((ULONG)(void *)p_hwnd, NULL, NULL, MAPI_LOGON_UI|MAPI_NO_MAIL|MAPI_EXTENDED, &lpMAPISession);
  MAPI_ASSERT_EX(hr);

  hr = lpMAPISession->OpenAddressBook(0,NULL,0,&lpAdrBook);
  MAPI_ASSERT_EX(hr);
  */

  //
  // Find DLL for Windows Address Book.
  HKEY wabKey;
  TCHAR szWABDllPath[MAX_PATH] = {0};
  DWORD keyValueByteCount = MAX_PATH * sizeof(TCHAR);
  if ( ERROR_SUCCESS == RegOpenKeyEx(HKEY_LOCAL_MACHINE, WAB_DLL_PATH_KEY, 0, KEY_READ, &wabKey) ) {
    if ( ERROR_SUCCESS != RegQueryValueEx(wabKey, "", 0, 0, (LPBYTE)szWABDllPath, &keyValueByteCount) ) {
      // In case of not found registry entry, use well known path.
      strcat(szWABDllPath, "C:\\Program Files\\Common Files\\System\\wab32");
    }
    RegCloseKey(wabKey);
  }

  //
  // Get pointer to WABOpen funtion.
  hInstLib = LoadLibrary(szWABDllPath);
  fWABOpen procWABOpen;
  if (hInstLib != NULL) {
    procWABOpen = (fWABOpen)GetProcAddress(hInstLib, "WABOpen"); 
    if (procWABOpen != NULL) {
      WAB_PARAM wp = {0};
      wp.cbSize = sizeof(WAB_PARAM);
      wp.hwnd = p_hwnd;
      if ( bEnableProfiles )
        wp.ulFlags = 0x00400000; // INFO: Unavailable under vs60: WAB_ENABLE_PROFILES
      hr = (procWABOpen)(&lpAdrBook, &lpWABObject, &wp, Reserved2);
      if (FAILED(hr)) return hr;
    } else return E_FAIL;
  } else return E_FAIL;

  bIsConnected = true;
    
  // Set default folder.
  hr = SetPABFolder();
  if ( FAILED(hr) )
    return hr;
  
  return hr;
}

HRESULT CWABWrapper::FreeBuffer(LPVOID lpBuffer) {
  HRESULT hr;
  if (lpWABObject)
    hr = lpWABObject->FreeBuffer(lpBuffer);
  else if (lpMAPISession)
    hr = MAPIFreeBuffer(lpBuffer);
  MAPI_ASSERT_EX(hr);
  return hr;
}

HRESULT CWABWrapper::Disconnect() {
  //delete cached categories.    
  ClearFolderMap();
  SAFE_RELEASE(lpAdrBook);
  SAFE_RELEASE(lpContainer);
  SAFE_RELEASE(lpWABObject);    
  SAFE_RELEASE(lpMAPISession);
  FreeLibrary(hInstLib);
  bIsConnected = false;
  MAPIUninitialize(); 
  return S_OK;
}

// --------------------------------------------------------

HRESULT CWABWrapper::GetFolders(_STL::vector<_STL::string>& aFolders) {
  HRESULT hr;
  IABContainer *pRootContainer = NULL;
  ULONG ulFlags = MAPI_BEST_ACCESS;
  ULONG ulObjType = NULL;
  LPUNKNOWN lpUnk = NULL;
  LPMAPITABLE pRootTable;
  LPSRowSet lpRows = NULL;

  // Get root container.
  hr = lpAdrBook->OpenEntry(0, 0, NULL, ulFlags, &ulObjType, &lpUnk );
  if (FAILED(hr)) return hr;
  pRootContainer = static_cast <IABContainer *>(lpUnk);
  
  hr = pRootContainer->GetHierarchyTable(CONVENIENT_DEPTH , &pRootTable);
  if (FAILED(hr)) {
    pRootContainer->Release();
    return hr;
  }
  
  static SizedSPropTagArray( 3, sptWabCatsCols ) = { 3, {PR_DEPTH, PR_ENTRYID, PR_DISPLAY_NAME}};
  hr = pRootTable->SetColumns((LPSPropTagArray)&sptWabCatsCols, 0);
  if (FAILED(hr)) {
    pRootTable->Release();
    pRootContainer->Release();
    return hr;
  }

  ULONG uRows = 0;
  hr = pRootTable->GetRowCount(0, &uRows);
  if (FAILED(hr)) {
    pRootTable->Release();
    pRootContainer->Release();
    return hr;
  }

  // Get all table rows (of specified columns).
  hr = pRootTable->QueryRows(uRows, 0, &lpRows);
  if (FAILED(hr)) {
    pRootTable->Release();
    pRootContainer->Release();
    return hr;
  }

  ClearFolderMap();

  // Now associate with each folder name its entryid.
  INT iCurDepth=0;
  mapi_TEntryid *pEntryId = NULL;
  for (UINT i=0; i<lpRows->cRows; i++){
    SRow *pRow = &lpRows->aRow[i];
    for (UINT v=0; v<pRow->cValues; v++) {        
      if (pRow->lpProps[v].ulPropTag == PR_ENTRYID) {
        if (pEntryId) delete pEntryId;
        pEntryId = new mapi_TEntryid(&pRow->lpProps[v]);
      } //PR_ENTRYID
      else
      if (pRow->lpProps[v].ulPropTag == PR_DISPLAY_NAME) {
        // Geto only first level forders.
        if (iCurDepth==1) {            
          _STL::string sCatName(pRow->lpProps[v].Value.lpszA);              
          WabFolders[sCatName] = pEntryId;
          pEntryId=NULL;
          aFolders.push_back(sCatName);
        } else {
          if (pEntryId) delete pEntryId;
          pEntryId = NULL;
        }
      } //PR_DISPLAY_NAME
      else if (pRow->lpProps[v].ulPropTag == PR_DEPTH){
        iCurDepth=pRow->lpProps[v].Value.l;           
      } //PR_DEPTH
    }
  }
  
  //
  FreeBuffer(lpRows);
  pRootTable->Release();
  pRootContainer->Release();
  return S_OK;
}

HRESULT CWABWrapper::GetCurFolderProps(_STL::string& sProps) {
  if ( lpContainer == NULL )
    return E_FAIL;
  LPSPropValue pprops = 0;
  ULONG ulCount = 0;
  HRESULT hr = lpContainer->GetProps(NULL, 0, &ulCount, &pprops);
  MAPI_ASSERT_EX(hr);

  sProps += _T("Properties for folder : ");
  sProps += GetCurFolderName();
  sProps += _T("\n");

  for ( INT i = 0; i < ulCount; ++i ) {
    const SPropValue& pv = pprops[i];
    sProps += GetFullPropertyString(pv);
    sProps += _T("\n");
  }

  hr = FreeBuffer(pprops);
  MAPI_ASSERT_EX(hr);
  return hr;
}

HRESULT CWABWrapper::SetFolder(const _STL::string& sFolder) {
  HRESULT hr;

  // First find folders EntryId, which was retrivied when calling GetFolders().
  mapi_TEntryid* pEntryId;
  FolderMap::iterator itr = WabFolders.find(sFolder);
  if (itr == WabFolders.end()) return false;
  pEntryId = itr->second;
  currentFolder = itr->second;

  if (lpContainer) lpContainer->Release();
  if (lpTable) lpTable->Release();

  // Open new container.
  ULONG ulObjType;
  IUnknown* lpUnk;
  hr = lpAdrBook->OpenEntry(pEntryId->size, (LPENTRYID)pEntryId->ab, NULL, 0, &ulObjType, &lpUnk );
  if (FAILED(hr)) {
    Connect();
    return hr;
  }
  lpContainer = static_cast <IABContainer *>(lpUnk);    
  
  // Get its contest table.
  hr = lpContainer->GetContentsTable(0x00100000/*WAB_LOCAL_CONTAINERS*/, &lpTable);
  if (hr == MAPI_E_NOT_FOUND) {
    // Might happend when folder is empty.    
    lpTable = NULL;
  } else if FAILED(hr) return hr;    

  // Filter out unwanted columns.
  if (lpTable) {
    if ( PropTags.Count() != 0 ) {
      hr = lpTable->SetColumns(PropTags, 0);
      if (FAILED(hr)) return hr;
    }
  }

  _STL::string sPropDump;
  GetCurFolderProps(sPropDump);
  OutputDebugString(sPropDump.c_str());
  
  return S_OK;
}

_STL::string CWABWrapper::GetCurFolderName() const { 

  for ( FolderMapConstItr itr = WabFolders.begin(); itr != WabFolders.end(); ++itr ) {

    // WAB Version
    if ( lpAdrBook != NULL )  {
      if ( currentFolder.isequal(lpAdrBook, *itr->second) )
        return itr->first;    
    }

    // MAPI Version
    else if ( lpMAPISession != NULL )  {
      if ( currentFolder.isequal(lpMAPISession, *itr->second) )
        return itr->first;    
    }

  }

  if ( lpAdrBook != NULL ) {
    // None of the subfolders, check PAB
    ULONG lpcbEntryID = 0;
    ENTRYID *lpEntryID = NULL;
    HRESULT hr = lpAdrBook->GetPAB(&lpcbEntryID, &lpEntryID);
    MAPI_ASSERT_EX(hr);
    if (SUCCEEDED(hr)) {
      mapi_TEntryid PABEntry(lpcbEntryID, lpEntryID);
      if ( PABEntry.isequal(lpAdrBook, PABEntry) )
        return _T("PAB");
    }
  }

  return _T("?"); 
}

HRESULT CWABWrapper::SetPABFolder() {  
  if ( !IsConnected() )
    return E_FAIL;
  
  //
  // Get entry identifier of default WAB container.
  ULONG lpcbEntryID = 0;
  ENTRYID *lpEntryID = NULL;
  HRESULT hr = lpAdrBook->GetPAB(&lpcbEntryID, &lpEntryID);
  if (FAILED(hr)) return hr;
  
  currentFolder.set(lpcbEntryID, lpEntryID);

  ULONG ulFlags = MAPI_BEST_ACCESS;
  ULONG ulObjType = NULL;
  LPUNKNOWN lpUnk = NULL;

  //
  // Get default folder containing contacts.
  hr = lpAdrBook->OpenEntry(lpcbEntryID, lpEntryID, NULL, ulFlags, &ulObjType, &lpUnk );
  if (FAILED(hr)) return hr;
  SAFE_RELEASE(lpContainer);
  lpContainer = static_cast <IABContainer *>(lpUnk);
  
  //
  // Now get ITable for all containers for current profile.
  hr = lpContainer->GetContentsTable(/*WAB_PROFILE_CONTENTS*/lpWABObject ? 0x00200000 : 0,&lpTable);
  if (FAILED(hr)) return hr;

  //
  // Set columns which we want to read during query operations.
  if ( Properties().Count() != 0 ) {
    hr = lpTable->SetColumns(PropTags, 0);
    if (FAILED(hr)) return hr;
  }
  
  return S_OK;
}

HRESULT CWABWrapper::CreateFolder(const _STL::string& sFolderName) {

  HRESULT      hr = S_OK;
  LPMAPISESSION   lpSession = NULL;
  LPADRBOOK      lpAddrbk = NULL;
  IABContainer   *lpAddrRoot = NULL;
  LPMAPITABLE      lpBooks = NULL;
  ULONG      ulCount = NULL;
  ULONG      ulObjType = NULL;
  LPSRowSet      pRows =   NULL;
  SRestriction           srName;
  SPropValue      spv;
  LPSPropValue    lpProp;
  LPABCONT      lpABC = NULL;
  LPMAPITABLE      lpTPLTable = NULL;
  LPMAPIPROP      lpNewEntry = NULL;

  SizedSPropTagArray(2, Columns) =
  {2, {PR_DISPLAY_NAME, PR_ENTRYID}};
  SizedSPropTagArray(2, TypeColumns) =
  {2, {PR_ADDRTYPE, PR_ENTRYID}};

  // Open root address book (container).
  hr = lpAdrBook->OpenEntry(0L,NULL,NULL,0L,&ulObjType,
    (LPUNKNOWN*)&lpAddrRoot);
  MAPI_ASSERT_EX(hr);

  // Get a table of all of the Address Books.
  hr = lpAddrRoot->GetHierarchyTable(CONVENIENT_DEPTH, &lpBooks);
  MAPI_ASSERT_EX(hr);
  //SHOWTABLE(lpBooks);

  // TODO: pomysl, znalesc "Tozsamosc glowna" i pod nia stworzyc nowy folder??

/*

  // Restrict the table to just its name and ID.
  hr = lpBooks->SetColumns((LPSPropTagArray)&Columns, 0);
  MAPI_ASSERT_EX(hr);

  // Build a restriction to find the Personal Address Book.
  srName.rt = RES_PROPERTY;
  srName.res.resProperty.relop = RELOP_EQ;
  srName.res.resProperty.ulPropTag = PR_DISPLAY_NAME;
  srName.res.resProperty.lpProp = &spv;
  spv.ulPropTag = PR_DISPLAY_NAME;
  spv.Value.lpszA = "Personal Address Book";  // Address Book to open

  // Apply the restriction
  hr = lpBooks->Restrict(&srName,0);
  MAPI_ASSERT_EX(hr);

  // Get the total number of rows returned. Typically, this will be 1.
  hr = lpBooks->GetRowCount(0,&ulCount);
  MAPI_ASSERT_EX(hr);

  // Get the row properties (trying to get the EntryID).
  hr = lpBooks->QueryRows(ulCount,0,&pRows);
  MAPI_ASSERT_EX(hr);

  // Get a pointer to the properties.
  lpProp = &pRows->aRow[0].lpProps[1];   // Point to the EntryID Prop

  // Open the Personal Address Book (PAB).
  hr = lpAddrbk->OpenEntry(lpProp->Value.bin.cb,
    (ENTRYID*)lpProp->Value.bin.lpb,
    NULL,MAPI_MODIFY,&ulObjType,
    (LPUNKNOWN FAR *)&lpABC);
  MAPI_ASSERT_EX(hr);

  // Get a table of templates for the address types.
  hr = lpABC->OpenProperty( PR_CREATE_TEMPLATES,
    (LPIID) &IID_IMAPITable,
    0, 0,
    (LPUNKNOWN *)&lpTPLTable);
  MAPI_ASSERT_EX(hr);
  SHOWTABLE(lpTPLTable);

  //// OLD CODE
*/

  // Ok, now you need to set some properties on the new Entry.
  const unsigned long cProps = 2;
  SPropValue   aPropsMesg[cProps];
  LPSPropProblemArray lpPropProblems = NULL;
  
  hr = /*lpContainer*/lpAddrRoot->CreateEntry(0, NULL, CREATE_CHECK_DUP_STRICT, &lpNewEntry);
  if(FAILED(hr))
    return E_FAIL;

  // INFO: look for article title : 'Folder and Address Book Container Properties' on MSDN
  //       for further reading on allowed folder properties.

  // Setup your properties.
  INT i = 0;
  aPropsMesg[i].dwAlignPad = 0;
  aPropsMesg[i].ulPropTag = PR_DISPLAY_NAME;
  aPropsMesg[i++].Value.LPSZ = (LPSTR)(LPCTSTR)sFolderName.c_str();

  //non wab
  //aPropsMesg[i].ulPropTag = PR_INSTANCE_KEY;
  //aPropsMesg[i].Value.bin.cb = sizeof(RECORD_KEY);
  //aPropsMesg[i++].Value.bin.lpb = (LPBYTE)&m_EntryID.rk; 
  
  aPropsMesg[i].dwAlignPad = 0;
  aPropsMesg[i].ulPropTag = PR_OBJECT_TYPE;
  aPropsMesg[i++].Value.ul = MAPI_ABCONT;

  //aPropsMesg[i].dwAlignPad = 0;
  //aPropsMesg[i].ulPropTag = PR_DISPLAY_TYPE;
  //aPropsMesg[i++].Value.l = DT_AGENT; 

  //aPropsMesg[i].dwAlignPad = 0;
  //aPropsMesg[i++].ulPropTag = PR_DEF_CREATE_MAILUSER;

 // aPropsMesg[i].ulPropTag = PR_FOLDER_SAVEASCAPONE;
 // aPropsMesg[i++].Value.b = FALSE;

 //aPropsMesg[i].ulPropTag = PR_HAS_RULES;
 //aPropsMesg[i++].Value.b = FALSE;

 //if (m_EntryID.rk.uid == ROOT_FOLDER_ID)
 //{
 //aPropsMesg[i].ulPropTag = PR_FOLDER_TYPE;
 //aPropsMesg[i++].Value.l = FOLDER_GENERIC;
 //}
 //else
 //{
 // aPropsMesg[i].ulPropTag = PR_FOLDER_TYPE;
 // aPropsMesg[i++].Value.l = FOLDER_GENERIC;
 //}

 //if (m_strMessageClass.GetLength() > 0)
 //{
 // aPropsMesg[i].ulPropTag = PR_DEFAULT_FORM;
 // aPropsMesg[i++].Value.lpszA = (LPSTR)(LPCTSTR)m_strMessageClass;
 //
 // aPropsMesg[i].ulPropTag = PR_MESSAGE_CLASS;
 // aPropsMesg[i++].Value.lpszA = (LPSTR)(LPCTSTR)m_strMessageClass;
 //}

 //if (m_strContainerClass.GetLength() > 0)
 //{
 // aPropsMesg[i].ulPropTag = PR_CONTAINER_CLASS;
 // aPropsMesg[i++].Value.lpszA = (LPSTR)(LPCTSTR)m_strContainerClass;
 //} 

  // Set the properties on the object.
  hr = lpNewEntry->SetProps(cProps, aPropsMesg, &lpPropProblems);  
  MAPI_ASSERT(hr);
  if(FAILED(hr))
    return E_FAIL;

  // Explictly save the changes to the new entry.
  hr = lpNewEntry->SaveChanges(0);
  MAPI_ASSERT(hr);
  //_STL::string s = GetMAPIError(hr);
  //OutputDebugString(s.c_str());

  // Cleanup
  if (lpNewEntry)
    lpNewEntry->Release();
  return hr;
} 

_STL::string CWABWrapper::GetMAPIError(HRESULT hr) {
  switch (hr) {
    case S_OK:                                      return "S_OK";
    case MAPI_E_BAD_CHARWIDTH:                      return "MAPI_E_BAD_CHARWIDTH: Either the MAPI_UNICODE flag was set and the implementation does not support Unicode, or MAPI_UNICODE was not set and the implementation only supports Unicode.";
    case MAPI_E_COMPUTED:                           return "MAPI_E_COMPUTED: The property cannot be updated because it is read-only, computed by the service provider responsible for the object.";
    case MAPI_E_INVALID_TYPE:                       return "MAPI_E_INVALID_TYPE: The property type is invalid.";
    case MAPI_E_NO_ACCESS:                          return "MAPI_E_NO_ACCESS: An attempt was made to perform an operation for which the user does not have permission. ";
    case MAPI_E_NOT_ENOUGH_MEMORY:                  return "MAPI_E_NOT_ENOUGH_MEMORY: Insufficient memory was available to perform this operation.";
    case MAPI_E_UNEXPECTED_TYPE:                    return "MAPI_E_UNEXPECTED_TYPE: The property type is not the type expected by the calling implementation.";
    case MAPI_E_NO_SUPPORT:                         return "MAPI_E_NO_SUPPORT: MAPI_E_NO_SUPPORT";
    case MAPI_E_OBJECT_CHANGED:                     return "MAPI_E_OBJECT_CHANGED: The object has changed since it was opened.";
    case MAPI_E_OBJECT_DELETED:                     return "MAPI_E_OBJECT_DELETED: The object has been deleted since it was opened.";

    case MAPI_E_CALL_FAILED:                        return "MAPI_E_CALL_FAILED";
    case MAPI_E_INVALID_PARAMETER:                  return "MAPI_E_INVALID_PARAMETER";
    case MAPI_E_INTERFACE_NOT_SUPPORTED:            return "MAPI_E_INTERFACE_NOT_SUPPORTED";
    case MAPI_E_STRING_TOO_LONG:                    return "MAPI_E_STRING_TOO_LONG";
    case MAPI_E_UNKNOWN_FLAGS:                      return "MAPI_E_UNKNOWN_FLAGS";
    case MAPI_E_INVALID_ENTRYID:                    return "MAPI_E_INVALID_ENTRYID";
    case MAPI_E_INVALID_OBJECT:                     return "MAPI_E_INVALID_OBJECT";
    case MAPI_E_BUSY:                               return "MAPI_E_BUSY";
    case MAPI_E_NOT_ENOUGH_DISK:                    return "MAPI_E_NOT_ENOUGH_DISK";
    case MAPI_E_NOT_ENOUGH_RESOURCES:               return "MAPI_E_NOT_ENOUGH_RESOURCES";
    case MAPI_E_NOT_FOUND :                         return "MAPI_E_NOT_FOUND ";
    case MAPI_E_VERSION:                            return "MAPI_E_VERSION";
    case MAPI_E_LOGON_FAILED:                       return "MAPI_E_LOGON_FAILED";
    case MAPI_E_SESSION_LIMIT:                      return "MAPI_E_SESSION_LIMIT";
    case MAPI_E_USER_CANCEL:                        return "MAPI_E_USER_CANCEL";
    case MAPI_E_UNABLE_TO_ABORT:                    return "MAPI_E_UNABLE_TO_ABORT";
    case MAPI_E_NETWORK_ERROR:                      return "MAPI_E_NETWORK_ERROR";
    case MAPI_E_DISK_ERROR:                         return "MAPI_E_DISK_ERROR";
    case MAPI_E_TOO_COMPLEX:                        return "MAPI_E_TOO_COMPLEX";
    case MAPI_E_BAD_COLUMN:                         return "MAPI_E_BAD_COLUMN";
    case MAPI_E_EXTENDED_ERROR:                     return "MAPI_E_EXTENDED_ERROR";
    case MAPI_E_CORRUPT_DATA:                       return "MAPI_E_CORRUPT_DATA";
    case MAPI_E_UNCONFIGURED:                       return "MAPI_E_UNCONFIGURED";
    case MAPI_E_FAILONEPROVIDER:                    return "MAPI_E_FAILONEPROVIDER";
    case MAPI_E_UNKNOWN_CPID:                       return "MAPI_E_UNKNOWN_CPID";
    case MAPI_E_UNKNOWN_LCID:                       return "MAPI_E_UNKNOWN_LCID";
    case MAPI_E_PASSWORD_CHANGE_REQUIRED:           return "MAPI_E_PASSWORD_CHANGE_REQUIRED";
    case MAPI_E_PASSWORD_EXPIRED:                   return "MAPI_E_PASSWORD_EXPIRED";
    case MAPI_E_INVALID_WORKSTATION_ACCOUNT:        return "MAPI_E_INVALID_WORKSTATION_ACCOUNT";
    case MAPI_E_INVALID_ACCESS_TIME:                return "MAPI_E_INVALID_ACCESS_TIME";
    case MAPI_E_ACCOUNT_DISABLED:                   return "MAPI_E_ACCOUNT_DISABLED";
    case MAPI_E_END_OF_SESSION:                     return "MAPI_E_END_OF_SESSION";
    case MAPI_E_UNKNOWN_ENTRYID:                    return "MAPI_E_UNKNOWN_ENTRYID";
    case MAPI_E_MISSING_REQUIRED_COLUMN:            return "MAPI_E_MISSING_REQUIRED_COLUMN";
    case MAPI_W_NO_SERVICE:                         return "MAPI_W_NO_SERVICE";
    case MAPI_E_BAD_VALUE:                          return "MAPI_E_BAD_VALUE";
    case MAPI_E_TYPE_NO_SUPPORT:                    return "MAPI_E_TYPE_NO_SUPPORT";
    case MAPI_E_TOO_BIG:                            return "MAPI_E_TOO_BIG";
    case MAPI_E_DECLINE_COPY:                       return "MAPI_E_DECLINE_COPY";
    case MAPI_E_UNEXPECTED_ID:                      return "MAPI_E_UNEXPECTED_ID";
    case MAPI_W_ERRORS_RETURNED:                    return "MAPI_W_ERRORS_RETURNED";
    case MAPI_E_UNABLE_TO_COMPLETE:                 return "MAPI_E_UNABLE_TO_COMPLETE";
    case MAPI_E_TIMEOUT:                            return "MAPI_E_TIMEOUT";
    case MAPI_E_TABLE_EMPTY:                        return "MAPI_E_TABLE_EMPTY";
    case MAPI_E_TABLE_TOO_BIG:                      return "MAPI_E_TABLE_TOO_BIG";
    case MAPI_E_INVALID_BOOKMARK:                   return "MAPI_E_INVALID_BOOKMARK";
    case MAPI_W_POSITION_CHANGED:                   return "MAPI_W_POSITION_CHANGED";
    case MAPI_W_APPROX_COUNT:                       return "MAPI_W_APPROX_COUNT";
    case MAPI_E_WAIT:                               return "MAPI_E_WAIT";
    case MAPI_E_CANCEL:                             return "MAPI_E_CANCEL";
    case MAPI_E_NOT_ME:                             return "MAPI_E_NOT_ME";
    case MAPI_W_CANCEL_MESSAGE:                     return "MAPI_W_CANCEL_MESSAGE";
    case MAPI_E_CORRUPT_STORE:                      return "MAPI_E_CORRUPT_STORE";
    case MAPI_E_NOT_IN_QUEUE:                       return "MAPI_E_NOT_IN_QUEUE";
    case MAPI_E_NO_SUPPRESS:                        return "MAPI_E_NO_SUPPRESS";
    case MAPI_E_COLLISION:                          return "MAPI_E_COLLISION";
    case MAPI_E_NOT_INITIALIZED:                    return "MAPI_E_NOT_INITIALIZED";
    case MAPI_E_NON_STANDARD:                       return "MAPI_E_NON_STANDARD";
    case MAPI_E_NO_RECIPIENTS:                      return "MAPI_E_NO_RECIPIENTS";
    case MAPI_E_SUBMITTED:                          return "MAPI_E_SUBMITTED";
    case MAPI_E_HAS_FOLDERS:                        return "MAPI_E_HAS_FOLDERS";
    case MAPI_E_HAS_MESSAGES:                       return "MAPI_E_HAS_MESSAGES";
    case MAPI_E_FOLDER_CYCLE:                       return "MAPI_E_FOLDER_CYCLE";
    case MAPI_W_PARTIAL_COMPLETION:                 return "MAPI_W_PARTIAL_COMPLETION";
    case MAPI_E_AMBIGUOUS_RECIP:                    return "MAPI_E_AMBIGUOUS_RECIP";
  };
  return "Unknown";
}

_STL::string CWABWrapper::GetFullPropertyString(const SPropValue& pv) const {  
  _STL::string sPropName = GetFullPropertyName(pv.ulPropTag);

  TCHAR	sBuffer[255];
  memset(sBuffer, 0, sizeof(sBuffer));
  ULONG propType = PROP_TYPE(pv.ulPropTag);
  
  if ( propType == PT_LONG ) {
    _sntprintf(sBuffer, sizeof(sBuffer)/sizeof(sBuffer[0]), _T("%.4X"), pv.Value.ul);
  }
  else if ( propType == PT_STRING8 ) {
    _sntprintf(sBuffer, sizeof(sBuffer)/sizeof(sBuffer[0]), _T("%s"), pv.Value.lpszA);
  }
  else if ( propType == PT_BINARY ) {
    _sntprintf(sBuffer, sizeof(sBuffer)/sizeof(sBuffer[0]), _T("bin;len=%d;"), pv.Value.bin.cb);
    int iStart = _tcslen(sBuffer);
    for ( INT i = 0; i < pv.Value.bin.cb && i < 25; ++i ) {
      _sntprintf(sBuffer+iStart+(i*3), sizeof(sBuffer)/sizeof(sBuffer[0])-iStart+i, _T(" %.2X"), pv.Value.bin.lpb[i]);  
    }
  }
  else {
    _sntprintf(sBuffer, sizeof(sBuffer)/sizeof(sBuffer[0]), _T("unk"));
  }

  sPropName += _T(", value=");
  sPropName += sBuffer;

  return sPropName;
}

_STL::string CWABWrapper::GetFullPropertyName(ULONG ulPropTag) const {
  _STL::string sPropName = GetPropertyName(ulPropTag);

  TCHAR	sBuffer[255];
  ULONG propId = PROP_ID(ulPropTag);
  ULONG propType = PROP_TYPE(ulPropTag);
  _sntprintf(sBuffer, sizeof(sBuffer)/sizeof(sBuffer[0]), _T("%-22s (PropID=%.4X, PropType=%.4X)"), sPropName.c_str(), propId, propType);
  return sBuffer;  
}

_STL::string CWABWrapper::GetPropertyName(ULONG ulPropTag) const {
  switch (ulPropTag) {
    case PROP_TAG( PT_LONG,0x3600):   return _T("PR_CONTAINER_FLAGS");
    case PROP_TAG( PT_LONG,0x3601):   return _T("PR_FOLDER_TYPE");
    case PROP_TAG( PT_LONG,0x3602):   return _T("PR_CONTENT_COUNT");
    case PROP_TAG( PT_LONG,0x3603):   return _T("PR_CONTENT_UNREAD");
    case PROP_TAG( PT_OBJECT,0x3604): return _T("PR_CREATE_TEMPLATES");
    case PROP_TAG( PT_OBJECT,0x3605): return _T("PR_DETAILS_TABLE");
    case PROP_TAG( PT_OBJECT,0x3607): return _T("PR_SEARCH");
    case PROP_TAG( PT_BOOLEAN,0x3609): return _T("PR_SELECTABLE");
    case PROP_TAG( PT_BOOLEAN,0x360A): return _T("PR_SUBFOLDERS");
    case PROP_TAG( PT_LONG,0x360B):    return _T("PR_STATUS");
    //case PROP_TAG( PT_TSTRING,0x360C): return _T("PR_ANR");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,0x360C): return _T("PR_ANR_W");
#else
    case PROP_TAG( PT_STRING8,0x360C): return _T("PR_ANR_A");
#endif
    case PROP_TAG( PT_MV_LONG,0x360D): return _T("PR_CONTENTS_SORT_ORDER");
    case PROP_TAG( PT_OBJECT,0x360E): return _T("PR_CONTAINER_HIERARCHY");
    case PROP_TAG( PT_OBJECT,0x360F): return _T("PR_CONTAINER_CONTENTS");
    case PROP_TAG( PT_OBJECT,0x3610): return _T("PR_FOLDER_ASSOCIATED_CONTENTS");
    case PROP_TAG( PT_BINARY,0x3611): return _T("PR_DEF_CREATE_DL");
    case PROP_TAG( PT_BINARY,0x3612): return _T("PR_DEF_CREATE_MAILUSER");
    //case PROP_TAG( PT_TSTRING,0x3613):  return _T("PR_CONTAINER_CLASS");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,0x3613):  return _T("PR_CONTAINER_CLASS_W");
#else
    case PROP_TAG( PT_STRING8,0x3613):  return _T("PR_CONTAINER_CLASS_A");
#endif
    case PROP_TAG( PT_I8,0x3614):       return _T("PR_CONTAINER_MODIFY_VERSION");
    case PROP_TAG( PT_BINARY,0x3615):   return _T("PR_AB_PROVIDER_ID");
    case PROP_TAG( PT_BINARY,0x3616):   return _T("PR_DEFAULT_VIEW_ENTRYID");
    case PROP_TAG( PT_LONG,0x3617):     return _T("PR_ASSOC_CONTENT_COUNT");

    case PROP_TAG( PT_BINARY,	0x0FFF): return _T("PR_ENTRYID");
    case PROP_TAG( PT_LONG,		0x0FFE): return _T("PR_OBJECT_TYPE");
    case PROP_TAG( PT_BINARY,	0x0FFD): return _T("PR_ICON");
    case PROP_TAG( PT_BINARY,	0x0FFC): return _T("PR_MINI_ICON");
    case PROP_TAG( PT_BINARY,	0x0FFB): return _T("PR_STORE_ENTRYID");
    case PROP_TAG( PT_BINARY,	0x0FFA): return _T("PR_STORE_RECORD_KEY");
    case PROP_TAG( PT_BINARY,	0x0FF9): return _T("PR_RECORD_KEY");
    case PROP_TAG( PT_BINARY,	0x0FF8): return _T("PR_MAPPING_SIGNATURE");
    case PROP_TAG( PT_LONG,		0x0FF7): return _T("PR_ACCESS_LEVEL");
    case PROP_TAG( PT_BINARY,	0x0FF6): return _T("PR_INSTANCE_KEY");
    case PROP_TAG( PT_LONG,		0x0FF5): return _T("PR_ROW_TYPE");
    case PROP_TAG( PT_LONG,		0x0FF4): return _T("PR_ACCESS");

    case PROP_TAG( PT_LONG,		0x3900): return _T("PR_DISPLAY_TYPE");
    case PROP_TAG( PT_BINARY,	0x3902): return _T("PR_TEMPLATEID");
    case PROP_TAG( PT_BINARY,	0x3904): return _T("PR_PRIMARY_CAPABILITY");

    case PROP_TAG( PT_LONG,		0x3000): return _T("PR_ROWID");

    //case PROP_TAG( PT_TSTRING,	0x3001): return _T("PR_DISPLAY_NAME");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,	0x3001): return _T("PR_DISPLAY_NAME_W");
#else
    case PROP_TAG( PT_STRING8,	0x3001): return _T("PR_DISPLAY_NAME_A");
#endif

    //case PROP_TAG( PT_TSTRING,	0x3002): return _T("PR_ADDRTYPE");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,	0x3002): return _T("PR_ADDRTYPE_W");
#else
    case PROP_TAG( PT_STRING8,	0x3002): return _T("PR_ADDRTYPE_A");
#endif

    //case PROP_TAG( PT_TSTRING,	0x3003): return _T("PR_EMAIL_ADDRESS");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,	0x3003): return _T("PR_EMAIL_ADDRESS_W");
#else
    case PROP_TAG( PT_STRING8,	0x3003): return _T("PR_EMAIL_ADDRESS_A");
#endif

    //case PROP_TAG( PT_TSTRING,	0x3004): return _T("PR_COMMENT");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,	0x3004): return _T("PR_COMMENT_W");
#else
    case PROP_TAG( PT_STRING8,	0x3004): return _T("PR_COMMENT_A");
#endif

    case PROP_TAG( PT_LONG,		0x3005): return _T("PR_DEPTH");

    //case PROP_TAG( PT_TSTRING,	0x3006): return _T("PR_PROVIDER_DISPLAY");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,	0x3006): return _T("PR_PROVIDER_DISPLAY_W");
#else
    case PROP_TAG( PT_STRING8,	0x3006): return _T("PR_PROVIDER_DISPLAY_A");
#endif

    case PROP_TAG( PT_SYSTIME,	0x3007): return _T("PR_CREATION_TIME");
    case PROP_TAG( PT_SYSTIME,	0x3008): return _T("PR_LAST_MODIFICATION_TIME");
    case PROP_TAG( PT_LONG,		0x3009): return _T("PR_RESOURCE_FLAGS");

    //case PROP_TAG( PT_TSTRING,	0x300A): return _T("PR_PROVIDER_DLL_NAME");
#if defined(UNICODE)
    case PROP_TAG( PT_UNICODE,	0x300A): return _T("PR_PROVIDER_DLL_NAME_W");
#else
    case PROP_TAG( PT_STRING8,	0x300A): return _T("PR_PROVIDER_DLL_NAME_A");
#endif

    case PROP_TAG( PT_BINARY,	0x300B): return _T("PR_SEARCH_KEY");
    case PROP_TAG( PT_BINARY,	0x300C): return _T("PR_PROVIDER_UID");
    case PROP_TAG( PT_LONG,		0x300D): return _T("PR_PROVIDER_ORDINAL");
  }
  return _T("UNKNOWN");
}

HRESULT CWABWrapper::SetFilter(_STL::string &sFilter, CDynamicPropertyTagArray& FilterProps) {

  if ( NULL == lpTable )
	  return E_FAIL;
  
  // Remove any current restriction
  HRESULT hr = lpTable->Restrict(NULL, 0);
  if ( FAILED(hr) )
    return hr;
  
  // If filter argument is empty then this call removes any current filtering.
  if (sFilter.empty()) {
    return S_OK;
  }

  // Specify search method, FL_SUBSTRING option does not work for WAB.
  const ULONG FUZZY_LEVEL_OPTS = FL_PREFIX | FL_IGNORECASE;

  aPropValue.resize(FilterProps.Count());
  aSResAnd.resize(FilterProps.Count());

  ZeroMemory(&SResRoot, sizeof(SResRoot));
  SResRoot.rt = RES_OR;
  SResRoot.res.resOr.cRes = FilterProps.Count();
  SResRoot.res.resOr.lpRes = &aSResAnd[0];

  for ( int i = 0; i < FilterProps.Count(); ++i ) {

    ULONG prop = FilterProps[i];

    // Skip non-searchable fields
    if ( prop == PR_OBJECT_TYPE || prop == PR_ENTRYID )
      continue;

    // Check if this is a multivalue property
    bool bMVI = false;
    if ( prop == PR_CONTACT_EMAIL_ADDRESSES )
      bMVI = true;

    //
    aPropValue[i].ulPropTag = prop;
    if ( bMVI )
      aPropValue[i].ulPropTag &= ~MV_FLAG;
    aPropValue[i].dwAlignPad = 0;
    aPropValue[i].Value.lpszA = (LPTSTR)sFilter.c_str();
    aSResAnd[i].rt = RES_CONTENT;
    aSResAnd[i].res.resContent.ulPropTag  = prop;
    aSResAnd[i].res.resContent.ulFuzzyLevel = FUZZY_LEVEL_OPTS;      
    aSResAnd[i].res.resContent.lpProp = &aPropValue[i];
  }
  
  hr = lpTable->Restrict(&SResRoot, 0 );
  return hr;
}

INT CWABWrapper::GetFolderItemCount() {
  ULONG iTotal=0;      
  HRESULT hr;
  if (lpTable == NULL) 
    return -1;
  else 
    hr = lpTable->GetRowCount(0, &iTotal);
  if (FAILED(hr))
    return -1;
  return iTotal;
}

// --------------------------------------------------------

int CWABWrapper::GetItemType(int iItemNum) {
  HRESULT hr;
  INT iCount = GetFolderItemCount();
  if (iItemNum < 0 || iItemNum >= iCount) return E_FAIL;

  // Only PR_OBJECT_TYPE is needed, but it wont work without the
  // other two.
  static SizedSPropTagArray( 3, sptObjectTypeCols ) = { 3, PR_OBJECT_TYPE, PR_ENTRYID, PR_DISPLAY_NAME };
  hr = lpTable->SetColumns(PropTags/*ObjectTypeCols*/, 0);
  if (FAILED(hr)) return hr;

  LONG iRowsSought;         
  hr = lpTable->SeekRow(((BOOKMARK) 0), iItemNum, &iRowsSought);
  if (FAILED(hr)) return hr;

  LPSRowSet lpRows = NULL;
  hr = lpTable->QueryRows(1, 0, &lpRows);
  if (FAILED(hr)) return hr;

  int iRet=E_FAIL;
  assert(lpRows->cRows!=0);
  if (lpRows->cRows!=0) {
    SRow *lpRow = &lpRows->aRow[0];      
    SPropValue *lpProp = &lpRow->lpProps[0];

    // Check if this is a distribution list or mail user (MAPI_MAILUSER).
    assert(lpProp->ulPropTag == PR_OBJECT_TYPE);
    iRet = (lpProp->Value.l == MAPI_DISTLIST) ? DIST_LIST : MAIL_USER;

    //FreeBuffer(lpRow);
    FreeBuffer(lpRows);
  }
  return iRet;
}

int CWABWrapper::GetItemType(mapi_TEntryid& ItemEntryID) {
  ULONG ulObjType = 0;  
  IUnknown* lpUnk;
    
  HRESULT hr = lpAdrBook->OpenEntry(ItemEntryID.size, (LPENTRYID)ItemEntryID.ab, 
    NULL, 0, &ulObjType, &lpUnk );
  if(FAILED(hr)) return hr;
  lpUnk->Release();

  if (ulObjType == MAPI_MAILUSER)
    return MAIL_USER;
  else 
    return DIST_LIST;
}
  
HRESULT CWABWrapper::GetDistList(int iItemNum, CDistList& DistList) {
  HRESULT hr;
  assert(lpAdrBook != NULL);

  INT iCount = GetFolderItemCount();
  if (iItemNum < 0 || iItemNum >= iCount) return E_FAIL;
  assert(GetItemType(iItemNum) == DIST_LIST);
      
  hr = lpTable->SetColumns(PropTags/*sptDistEntryId*/, 0);
  assert(SUCCEEDED(hr));
  if (FAILED(hr)) return hr;
  
  // Seek to the row with distribution data of interest.
  LONG iRowsSought;
  hr = lpTable->SeekRow(((BOOKMARK) 0), iItemNum, &iRowsSought);
  assert(SUCCEEDED(hr));
  if (FAILED(hr)) return hr;

  // Get one row and fill CDistList object with its data (no emails are loaded here).
  LPSRowSet lpRows = NULL;
  hr = lpTable->QueryRows(1, 0, &lpRows);
  assert(SUCCEEDED(hr));
  if (FAILED(hr)) return hr;

  assert(lpRows->cRows==1 /*&& lpRows->aRow[0].cValues == 3*/ && lpRows->aRow[0].lpProps[1].ulPropTag == PR_ENTRYID);
  mapi_TEntryid DistEntryId = &lpRows->aRow[0].lpProps[1];

  DistList.FillData(&lpRows->aRow[0]);          

  FreeBuffer(&lpRows->aRow[0]);
  FreeBuffer(lpRows); 

  // Now emails will be loaded.
  ULONG ulObjType;
  IUnknown* lpUnk;
  hr = lpAdrBook->OpenEntry(DistEntryId.size, (LPENTRYID)DistEntryId.ab, NULL, 0, &ulObjType, &lpUnk );
  if (FAILED(hr)) {
    assert(false); return hr;
  }
  
  IABContainer* lpDLContainer = static_cast <IABContainer *>(lpUnk);  
  assert(ulObjType == MAPI_DISTLIST);
  if (ulObjType!=MAPI_DISTLIST) {
    lpDLContainer->Release(); return E_FAIL;
  }
    
  // Get table with all members of this distribution list.
  IMAPITable *pDLTable;
  hr = lpDLContainer->GetContentsTable(0, &pDLTable);
  if (FAILED(hr)) {
    lpDLContainer->Release(); return E_FAIL;
  }

  hr = pDLTable->SetColumns((LPSPropTagArray) &sptDistEntryIdCols, 0);
  if (FAILED(hr)) {
    pDLTable->Release(); lpDLContainer->Release(); 
    return E_FAIL;
  }

  DWORD dwMemberCount;
  hr = pDLTable->GetRowCount(0, &dwMemberCount);
  if (FAILED(hr))
    return E_FAIL;

  if (dwMemberCount != 0) {
    lpRows = NULL;
    hr = lpTable->QueryRows(dwMemberCount, 0, &lpRows);
    
    // Iterate each row and copy found ENTRYID to CDistList objects array for later CMailUser retrieval.
    assert(lpRows->cRows == dwMemberCount);
    for (int k=0; k < lpRows->cRows; ++k) {
      DistList.AddMember(mapi_TEntryid());
      assert(/*lpRows->aRow[k].cValues == 3 && */lpRows->aRow[0].lpProps[1].ulPropTag == PR_ENTRYID);
      DistList.GetMember(DistList.MembersCount() - 1) = &lpRows->aRow[k].lpProps[1];   
      FreeBuffer(&lpRows->aRow[k]);
    }
    FreeBuffer(lpRows);    
  }
  pDLTable->Release();
  lpDLContainer->Release();
 
  return S_OK;
}

HRESULT CWABWrapper::GetMailUser(int iItemNum, CMailUser& MailUser) {
  HRESULT hr;
  INT iCount = GetFolderItemCount();
  if (iItemNum < 0 || iItemNum >= iCount) return E_FAIL;
  assert(GetItemType(iItemNum) == MAIL_USER);

  if (iItemNum >= iCount) return E_FAIL;
  LONG iRowsSought;         
  
  hr = lpTable->SetColumns(PropTags, 0);
  assert(SUCCEEDED(hr));
  if (FAILED(hr)) return hr;
  
  hr = lpTable->SeekRow(((BOOKMARK) 0), iItemNum, &iRowsSought);
  assert(SUCCEEDED(hr));
  if (FAILED(hr)) return hr;

  // Pobierz jeden wiersz 
  LPSRowSet lpRows = NULL;
  hr = lpTable->QueryRows(1, 0, &lpRows);
  assert(SUCCEEDED(hr));
  if (FAILED(hr)) return hr;

  if (lpRows->cRows!=0) {
    SRow *lpRow = &lpRows->aRow[0];      
    MailUser.FillData(lpRow);          
    FreeBuffer(lpRows);    
  }
  return S_OK;
}

HRESULT CWABWrapper::UpdateMailUser(CMailUser& MailUser) {
  ULONG ulObjType = 0;  
  IUnknown* lpUnk;
    
  HRESULT hr = lpAdrBook->OpenEntry(MailUser.GetEntryID().size, (LPENTRYID)MailUser.GetEntryID().ab, 
    NULL, MAPI_BEST_ACCESS|MAPI_DEFERRED_ERRORS, &ulObjType, &lpUnk );
  assert(ulObjType == MAPI_MAILUSER);
  if(FAILED(hr)) return hr;
  IABContainer* lpMUContainer = static_cast <IABContainer *>(lpUnk);
  SPropTagArray Props;
  SRow Row;

  SPropValue prop;
  prop.dwAlignPad = 0;
  prop.ulPropTag = PR_DISPLAY_NAME;
  prop.Value.LPSZ = (LPSTR)(LPCTSTR)MailUser.GetStringPropValue(PR_DISPLAY_NAME).c_str();

  hr = lpMUContainer->SetProps(1, &prop, NULL);
  if (FAILED(hr)) return hr;

  hr = lpMUContainer->SaveChanges(NULL); 
  if (FAILED(hr)) return hr;

  lpMUContainer->Release();
  return hr;
}

HRESULT CWABWrapper::GetMailUser(mapi_TEntryid& MailUserEntryID, CMailUser& MailUser) {
  ULONG ulObjType = 0;  
  IUnknown* lpUnk;
    
  HRESULT hr = lpAdrBook->OpenEntry(MailUserEntryID.size, (LPENTRYID)MailUserEntryID.ab, 
    NULL, 0, &ulObjType, &lpUnk );
  assert(ulObjType == MAPI_MAILUSER);
  if(FAILED(hr)) return hr;
  IABContainer* lpMUContainer = static_cast <IABContainer *>(lpUnk);
  SPropTagArray Props;
  SRow Row;
  hr = lpMUContainer->GetProps(PropTags, 0, &Row.cValues, &Row.lpProps);
  if (FAILED(hr)) return hr;

  MailUser.FillData(&Row);
  lpMUContainer->Release();
  return hr;
}

static SizedSPropTagArray(1, tagaDefaultDL) = {1, {PR_DEF_CREATE_DL}};
HRESULT CWABWrapper::UpdateDistList(CDistList& DistList) {
  HRESULT hr;

  LPDISTLIST pDistList=NULL;
  // If there is no valid entry id then treat this as a new distribution list.
  if (DistList.GetEntryID().isempty()) {
    AddDistList(DistList);
  } else {
    //this->lpAdrBook->

    //  LPMAILUSER pMailUser=NULL;
    //PR_DEF_CREATE_MAILUSER  
  }
  return E_NOTIMPL;
}

HRESULT CWABWrapper::AddDistList(CDistList& DistList) {
  // Create new distribution list.
  HRESULT hr;
  ULONG lpcbEntryID;
  LPDISTLIST pDistList=NULL;
  ENTRYID *lpEntryID;
  hr = lpAdrBook->GetPAB(&lpcbEntryID, &lpEntryID);
  if (FAILED(hr)) return hr;

  ULONG ulObjType = 0;
  LPABCONT lpPABCont = NULL;
  IUnknown* lpUnk;
  hr = lpAdrBook->OpenEntry(lpcbEntryID, lpEntryID,
    NULL,
    MAPI_MODIFY,
    &ulObjType,
    &lpUnk);
  FreeBuffer(lpEntryID);  
  if (FAILED(hr)) return hr;
  lpPABCont = (LPABCONT)lpUnk;

  LPSPropValue lpspvDefDLTpl;
  ULONG ulcb = 0;
  SPropTagArray props;
  props.aulPropTag[0]=PR_DEF_CREATE_DL;
  props.cValues = 1;
  hr = lpPABCont->GetProps(&props, 0,
    &ulcb,
    &lpspvDefDLTpl);
  if (FAILED(hr)) return hr;

  LPMAPIPROP lpNewPABEntry;
  hr = lpPABCont->CreateEntry(lpspvDefDLTpl->Value.bin.cb,
    (LPENTRYID)lpspvDefDLTpl->Value.bin.lpb, CREATE_CHECK_DUP_STRICT, &lpNewPABEntry);
  if (FAILED(hr)) return hr;

  hr = lpNewPABEntry->QueryInterface(IID_IDistList, (void**)&pDistList);
  if (FAILED(hr)) return hr;

  // TODO: przeniesc dodawanie danych jako zewnetrna funkcja.
  SPropValue PropVal;
  PropVal.dwAlignPad = 0;
  PropVal.ulPropTag = PR_DISPLAY_NAME;
  TCHAR *acBuffer = (TCHAR*)_alloca(DistList.GetStringPropValue(PR_DISPLAY_NAME).size() + 1);
  strcpy(acBuffer, DistList.GetStringPropValue(PR_DISPLAY_NAME).c_str());
  PropVal.Value.lpszA = acBuffer;
  LPSPropProblemArray lpProbArray;
  hr = lpNewPABEntry->SetProps(1, &PropVal, &lpProbArray);
  if (FAILED(hr)) return hr;

  hr = lpNewPABEntry->SaveChanges(KEEP_OPEN_READWRITE);
  if (FAILED(hr)) return hr;

  pDistList->Release();
  FreeBuffer(lpspvDefDLTpl);
  return E_NOTIMPL;
}

HRESULT CWABWrapper::AddMailUser(CMailUser& MailUser) {
  // Create new mail user.
  HRESULT hr;
  ULONG lpcbEntryID;
  LPMAILUSER pMailUser=NULL;
  ENTRYID *lpEntryID;
  hr = lpAdrBook->GetPAB(&lpcbEntryID, &lpEntryID);
  if (FAILED(hr)) return hr;

  ULONG ulObjType = 0;
  LPABCONT lpPABCont = NULL;
  IUnknown* lpUnk;
  hr = lpAdrBook->OpenEntry(lpcbEntryID, lpEntryID,
    NULL,
    MAPI_MODIFY,
    &ulObjType,
    &lpUnk);
  FreeBuffer(lpEntryID);  
  if (FAILED(hr)) return hr;
  lpPABCont = (LPABCONT)lpUnk;

  LPSPropValue lpspvDefMUTpl;
  ULONG ulcb = 0;
  SPropTagArray props;
  props.aulPropTag[0]=PR_DEF_CREATE_MAILUSER;
  props.cValues = 1;
  hr = lpPABCont->GetProps(&props, 0,
    &ulcb,
    &lpspvDefMUTpl);
  if (FAILED(hr)) return hr;

  LPMAPIPROP lpNewPABEntry;
  hr = lpPABCont->CreateEntry(lpspvDefMUTpl->Value.bin.cb,
    (LPENTRYID)lpspvDefMUTpl->Value.bin.lpb, CREATE_CHECK_DUP_STRICT, &lpNewPABEntry);
  if (FAILED(hr)) return hr;

  hr = lpNewPABEntry->QueryInterface(IID_IMailUser, (void**)&pMailUser);
  if (FAILED(hr)) return hr;

  // 
  SPropValue PropVal;
  //hr = MAPIAllocateBuffer(sizeof(SPropValue) * cProps, (LPVOID *) &pPropVal);
  //if (FAILED(hr))
  //  return hr;
  
  if ( MailUser.GetStringPropValue(PR_DISPLAY_NAME).size() ) {
    PropVal.dwAlignPad = 0;  
    PropVal.ulPropTag = PR_DISPLAY_NAME;
    TCHAR *acBuffer = (TCHAR*)_alloca(MailUser.GetStringPropValue(PR_DISPLAY_NAME).size() + 1);
    strcpy(acBuffer, MailUser.GetStringPropValue(PR_DISPLAY_NAME).c_str());
    PropVal.Value.lpszA = acBuffer;

    LPSPropProblemArray lpProbArray = NULL;
    hr = pMailUser->SetProps(1, &PropVal, &lpProbArray);
    if (FAILED(hr)) return hr;

    hr = lpNewPABEntry->SaveChanges(FORCE_SAVE);
    if (FAILED(hr)) return hr;
  }

  if ( MailUser.GetStringPropValue(PR_GIVEN_NAME).size() ) {
    PropVal.dwAlignPad = 0;  
    PropVal.ulPropTag = PR_GIVEN_NAME;
    TCHAR *acBuffer = (TCHAR*)_alloca(MailUser.GetStringPropValue(PR_GIVEN_NAME).size() + 1);
    strcpy(acBuffer, MailUser.GetStringPropValue(PR_GIVEN_NAME).c_str());
    PropVal.Value.lpszA = acBuffer;

    LPSPropProblemArray lpProbArray = NULL;
    hr = pMailUser->SetProps(1, &PropVal, &lpProbArray);
    if (FAILED(hr)) return hr;

    hr = lpNewPABEntry->SaveChanges(FORCE_SAVE);
    if (FAILED(hr)) return hr;
  }

  //MAPIFreeBuffer((LPVOID)pPropVal); 
  pMailUser->Release();
  FreeBuffer(lpspvDefMUTpl);
  return S_OK;
}

//
// CPropBase
//

HRESULT CPropBase::FillData(const SRow *pData) {

  PropValues.clear();

  for(ULONG j=0;j<pData->cValues;j++) {
    SPropValue *lpProp = &pData->lpProps[j];

    if (lpProp->ulPropTag == PR_ENTRYID)
      EntryID = lpProp;

    if (PROP_TYPE(lpProp->ulPropTag) == PT_TSTRING) {
      PropValues[lpProp->ulPropTag] = lpProp->Value.LPSZ;
    }

    // Parse multivalue e-mail property.
    if (PROP_TYPE(lpProp->ulPropTag) == PT_MV_TSTRING) {
      INT iCount = lpProp->Value.MVszA.cValues;
      _STL::vector<_STL::string> aValues;
      if (lpProp->Value.MVszA.lppszA) {
        for (INT i = 0; i<iCount; ++i) {
          aValues.push_back(lpProp->Value.MVszA.lppszA[i]);
        }
      }
      AddStringPropMultiValue(lpProp->ulPropTag, aValues);
    }    
  }  
  
  assert(EntryID.size != 0);

  return S_OK;
}
