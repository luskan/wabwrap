#ifndef _WABWRAPPER_H
#define _WABWRAPPER_H

//#define MIDL_PASS
#include <mapix.h>
#include <wab.h>  
//#undef MIDL_PASS

#include <string>
#include <vector>
#include <map>
#include <utility>
#include <tchar.h>

#include "MAPIASST.H" // from kb177542

#define _STL std

//
// Very usefull class from Lucian's Wischnik mapi utils which can be found here:
// http://www.wischik.com/lu/programmer/mapi_utils.html.
class mapi_TEntryid
{ public:
  unsigned int size;
  ENTRYID *ab;
  mapi_TEntryid() {ab=0;size=0;}
  mapi_TEntryid(SPropValue *v) {ab=0;size=0; if (v->ulPropTag!=PR_ENTRYID) return; set(v->Value.bin.cb,(ENTRYID*)v->Value.bin.lpb);}
  mapi_TEntryid(const mapi_TEntryid &e) {ab=0;size=0; set(e.size,e.ab);}
  mapi_TEntryid(unsigned int asize,ENTRYID *eid) {ab=0;size=0; set(asize,eid);}
  mapi_TEntryid &operator= (const mapi_TEntryid& e) {set(e.size,e.ab); return *this;}
  mapi_TEntryid &operator= (const mapi_TEntryid *e) {set(e->size,e->ab); return *this;}
  mapi_TEntryid &operator= (const SPropValue *v) {set(0,0); if (PROP_TYPE(v->ulPropTag)!=PT_BINARY) return *this; set(v->Value.bin.cb,(ENTRYID*)v->Value.bin.lpb); return *this;}
  ~mapi_TEntryid() {set(0,0);}
  void set(unsigned int asize, ENTRYID *eid) {if (ab!=0) { delete[] ((char*)ab); } size = asize; if (eid == 0) ab=0; else { ab=(ENTRYID*)(new char[size]); memcpy(ab, eid,size);} }
  void clear() {set(0,0);}
  bool isempty() const {return (ab==0 || size==0);}
  bool isequal(IMAPISession *sesh, mapi_TEntryid const &e) const
  { if (isempty() || e.isempty()) return false;
    ULONG res; HRESULT hr = sesh->CompareEntryIDs(size,ab,e.size,e.ab,0,&res);
    if (hr!=S_OK) return false;
    return (res!=0);
  }
  bool isequal(LPADRBOOK lpAdrBook, mapi_TEntryid const &e) const
  { if (isempty() || e.isempty()) return false;
    ULONG res; HRESULT hr = lpAdrBook->CompareEntryIDs(size,ab,e.size,e.ab,0,&res);
    if (hr!=S_OK) return false;
    return (res!=0);
  }
};

//
// Wrapper class to Property Tag-s.
// It makes it easier to manage SPropTagArray structure. 
class CDynamicPropertyTagArray {
public:
  CDynamicPropertyTagArray() { pProps = NULL; }
  ~CDynamicPropertyTagArray() { Invalidate(); }

  // Adds new property tag value. Use values from mapitags.h : PR_*
  void Add(ULONG ulProp) {
    Invalidate();
    aPropTags.push_back(ulProp);
  }

  // Returns number of added property tag values
  int Count() {
    return (int)aPropTags.size();
  }

  // Clears whole structure
  void Clear() {
    Invalidate();
    aPropTags.clear();
  }

  // Returns index-th property tag value.
  ULONG operator[](int index) {
    return aPropTags[index];
  }

  // Implicitly casts to pointer to SPropTagArray structure, which is required
  // by many MAPI functions.
  operator SPropTagArray*() {
    if ( NULL == pProps ) {
      pProps = (SPropTagArray*)new char[CbNewSPropTagArray(aPropTags.size())];
      pProps->cValues = (ULONG)aPropTags.size();
      for ( int i = 0; i < aPropTags.size(); ++i )
        pProps->aulPropTag[i] = aPropTags[i];        
    }
    return pProps;
  }
  
private:
  void Invalidate() { 
    if ( pProps ) delete[] pProps; 
    pProps = NULL; 
  }
  SPropTagArray* pProps;
  _STL::vector<ULONG> aPropTags;
};

class CPropBase {
  
public:

  // Returns Entry ID
  mapi_TEntryid& GetEntryID() { return EntryID; }

  // Returns value for given property tag, or empty _STL::string if no such tag exists.
  _STL::string GetStringPropValue(ULONG PropValue) {
    _STL::map<ULONG, _STL::string>::iterator itr = PropValues.find(PropValue);
    if ( itr == PropValues.end() )
      return "";
    else
      return itr->second;
  }

  // Adds new property _STL::string tag value (or updates previous value).
  HRESULT AddStringPropValue(ULONG PropValue, const _STL::string& Value) {
    if ( PROP_TYPE(PropValue) != PT_TSTRING )
      return E_FAIL;
    PropValues[PropValue] = Value;
    return S_OK;
  }

  // Returns true if given property tag exists.
  bool PropStringTagExists(ULONG PropValue) {
    _STL::map<ULONG, _STL::string>::iterator itr = PropValues.find(PropValue);
    return itr != PropValues.end();    
  }

  // Adds new multi value _STL::string property tag (or updates previous value).
  HRESULT AddStringPropMultiValue(ULONG PropValue, const _STL::vector<_STL::string>& aValues) {
    if ( PROP_TYPE(PropValue) != PT_MV_TSTRING )
      return E_FAIL;
    PropMultiValues[PropValue] = aValues;
    return S_OK;
  }

  // Return multi value _STL::string property tag.
  HRESULT GetStringPropMultiValue(ULONG PropValue, _STL::vector<_STL::string>& aValues) {
    if ( PROP_TYPE(PropValue) != PT_MV_TSTRING )
      return E_FAIL;
    aValues = PropMultiValues[PropValue];
    return S_OK;
  }

  // Returns true if given multi value property tag exists.
  bool PropMultiValueStringTagExists(ULONG PropValue) {
    _STL::map<ULONG, _STL::vector<_STL::string> >::iterator itr = PropMultiValues.find(PropValue);
    return itr != PropMultiValues.end();    
  }

  // Initializes this instance with property values.
  HRESULT FillData(const SRow *pData);

private:

  // Property Entry ID.
  mapi_TEntryid EntryID;

  // Property values, key is a property tag like PR_DISPLAY_NAME, PR_HOME_ADDRESS_STREET, ...
  // Value is a value for giver property tag.
  _STL::map<ULONG, _STL::string> PropValues;

  // Map contains _STL::string multi values.
  _STL::map<ULONG, _STL::vector<_STL::string> > PropMultiValues;

  // Entry IDs to mail users that are members of this distribution list.
  _STL::vector<mapi_TEntryid> aMembers;
};

///
/// Class that wraps distribution lists.
class CDistList : public CPropBase {  
 
public:
   
  // Adds new distribution list member
  void AddMember(const mapi_TEntryid& MailUserEID) { 
    aMembers.push_back(MailUserEID); 
  }

  // Returns reference to ind-th member
  mapi_TEntryid& GetMember(int ind) { return aMembers[ind]; }

  // Returns number of members
  size_t MembersCount() { return aMembers.size(); }

private:

  // Entry IDs to mail users that are members of this distribution list.
  _STL::vector<mapi_TEntryid> aMembers;
};

///
/// Mail user-s are members of distribution lists. They can have 
/// multiple different email addresses.
class CMailUser : public CPropBase {
};

///
/// Wrapper for basic Windows Address Book operations.
class CWABWrapper {
  LPMAPISESSION lpMAPISession;  // Used when connection is being done using MAPILogonEx
  LPWABOBJECT lpWABObject;      // Used when connection is being done using WABOpen.

  HINSTANCE hInstLib;
  LPADRBOOK lpAdrBook;
  IABContainer *lpContainer;
  ULONG lpcbEntryID;
  ENTRYID *lpEntryID;
  LPMAPITABLE lpTable;
  
  typedef _STL::map<_STL::string, mapi_TEntryid*> FolderMap;
  typedef FolderMap::iterator FolderMapItr;
  typedef FolderMap::const_iterator FolderMapConstItr;
  FolderMap WabFolders;

  mapi_TEntryid currentFolder; ///< EntryID for currently selected folder.

  // Properties which should be accessible.
  CDynamicPropertyTagArray PropTags;

  bool bIsConnected;      // Connected/disconnected state.

  // Used by restriction setter.
  SRestriction SResRoot;
  _STL::vector<SPropValue> aPropValue;
  _STL::vector<SRestriction> aSResAnd;

  //
  // Helpers.
  void ClearFolderMap();
  HRESULT Remove(mapi_TEntryid& EntryID);
  
public:

  CWABWrapper();
  ~CWABWrapper();

  //
  // Initialization / deinitialiation functions.

  /// Connects to Windows Address Book, and sets up default container (folder).
  /// If there was already connection, it will be disconnected and new 
  /// conection will be made.
  /// \param bEnableProfiles if true then it will be possible to choose identities folders,
  ///     if false, then no multiuser functionality is present.
  ///     To read more on this topic read in MSDN article: 'WAB and Multi-User/Multi-Identity Profiles'.
  HRESULT Connect(HWND p_hwnd = NULL, bool bEnableProfiles = true);
  LPADRBOOK GetAdrBook() { return lpAdrBook; }
  HRESULT FreeBuffer(LPVOID lpBuffer);

  /// Disconnects, clears all tables/variables.
  HRESULT Disconnect();

  /// Returns whether there was a successful connection done.
  bool IsConnected() { return bIsConnected; }

  //
  // Properties management functions
  CDynamicPropertyTagArray& Properties() { return PropTags; }

  //
  // Folder management functions

  /// Returns array of folder names.
  /// \param aFolders OUT parameter that will contain folder names.
  HRESULT GetFolders(_STL::vector<_STL::string>& aFolders);

  /// Sets active folder to [sFolder].
  /// \param Name of the folder which to make active. Pass empty _STL::string to use global folder.
  HRESULT SetFolder(const _STL::string& sFolder);
  _STL::string GetCurFolderName() const;
  
  _STL::string GetPropertyName(ULONG ulPropTag) const;
  _STL::string GetFullPropertyName(ULONG ulPropTag) const;
  _STL::string GetFullPropertyString(const SPropValue& pv) const;

  HRESULT GetCurFolderProps(_STL::string& sProps);

  /// Sets default folder.
  HRESULT SetPABFolder();

  /// Creates subfolder in currently selected folder.
  HRESULT CreateFolder(const _STL::string& sFolderName);

  /// Returns _STL::string representation of MAPI error code.
  static _STL::string GetMAPIError(HRESULT hr);

  /// Filters currently selected folder with specified _STL::string.
  HRESULT SetFilter(_STL::string &sFilter) { return SetFilter(sFilter, PropTags); }
  HRESULT SetFilter(_STL::string &sFilter, CDynamicPropertyTagArray& FilterProps);  

  /// Returns number of item count in filtered or not filtered folder.
  INT GetFolderItemCount();
    
  //
  // Item type verification.
  static const int MAIL_USER;
  static const int DIST_LIST;
  int GetItemType(int iItemNum);
  int GetItemType(mapi_TEntryid& ItemEntryID);

  //
  // Distribution list management.
  HRESULT GetDistList(int iItemNum, CDistList& DistList);  
  HRESULT UpdateDistList(CDistList& DistList);
  HRESULT AddDistList(CDistList& DistList);
  HRESULT RemoveDistList(CDistList& DistList);
  
  // 
  // Mail user management.  
  HRESULT GetMailUser(int iItemNum, CMailUser& MailUser);
  HRESULT GetMailUser(mapi_TEntryid& MailUserEntryID, CMailUser& MailUser);  
  HRESULT UpdateMailUser(CMailUser& MailUser);
  HRESULT AddMailUser(CMailUser& MailUser);
  HRESULT RemoveMailUser(CMailUser& MailUser); 
};

#endif