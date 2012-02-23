// Test.cpp : Defines the entry point for the console application.
//

#include <stdio.h>
#include <assert.h>
#include <CRTDBG.h>
#include "..\WABWrapper.h"

#define TEST() printf("TESTING: %s\n", __FUNCTION__)  
#define ASSERT(a) AssertTrue(a, "@"##__FUNCTION__##",\n expr: "#a, __LINE__)
void AssertTrue(bool b, char* sText, INT iLine) {
  if ( !b ) {
    printf("ERR: \n %s\n line: %d\n", sText, iLine);
    exit(-1);//_CrtDbgBreak();
  }
  //assert(b);
}

class CDynamicPropertyTagArrayTestCase {
public:
  static void AddTestCase() {
    TEST();
    CDynamicPropertyTagArray PropTagArr;
    PropTagArr.Add(PR_OBJECT_TYPE);
    PropTagArr.Add(PR_ENTRYID);
    PropTagArr.Add(PR_DISPLAY_NAME);
    ASSERT(PropTagArr.Count() == 3);
  }

  static void CountTestCase() {
    TEST();
    CDynamicPropertyTagArray PropTagArr;
    PropTagArr.Add(PR_OBJECT_TYPE);
    PropTagArr.Add(PR_ENTRYID);    
    ASSERT(PropTagArr.Count() == 2);
    PropTagArr.Add(PR_GIVEN_NAME);
    PropTagArr.Add(PR_MIDDLE_NAME);    
    ASSERT(PropTagArr.Count() == 4);    
  }

  static void ClearTestCase() {
    TEST();
    CDynamicPropertyTagArray PropTagArr;
    for ( INT k = 0; k < 10; ++k ) {
      for ( INT p = 0; p < k*2; ++p ) {
        PropTagArr.Add(PR_OBJECT_TYPE);
        PropTagArr.Add(PR_ENTRYID);    
        PropTagArr.Add(PR_GIVEN_NAME);
        PropTagArr.Add(PR_MIDDLE_NAME);    
      }
      ASSERT(PropTagArr.Count() == k*2*4);
      PropTagArr.Clear();
      ASSERT(PropTagArr.Count() == 0);
    }
  }

  static void LPSPropTagArrayOpTestCase() {
    TEST();
    CDynamicPropertyTagArray PropTagArr;
    PropTagArr.Add(PR_OBJECT_TYPE);
    PropTagArr.Add(PR_ENTRYID);        
    ASSERT(((LPSPropTagArray)PropTagArr)->cValues == 2);
    ASSERT(((LPSPropTagArray)PropTagArr)->aulPropTag[0] == PR_OBJECT_TYPE);
    ASSERT(((LPSPropTagArray)PropTagArr)->aulPropTag[1] == PR_ENTRYID);
  }

  static void IndexOpTestCase() {
    TEST();
    CDynamicPropertyTagArray PropTagArr;
    PropTagArr.Add(PR_OBJECT_TYPE);
    PropTagArr.Add(PR_ENTRYID);            
    ASSERT(PropTagArr[0] == PR_OBJECT_TYPE);
    ASSERT(PropTagArr[1] == PR_ENTRYID);
  }
};

class CPropBaseTestCase {
public:
  static void GetEntryIDTestCase() {    
    TEST();
    CPropBase PropBase;
    ASSERT(PropBase.GetEntryID().isempty());
    
    // Create sample row set.
    SizedSRowSet(1, TestRow) = {1, { NULL, NULL, NULL } };
    TestRow.aRow[0].cValues = 1;

    SPropValue PropValue[1] = { NULL };
    PropValue[0].ulPropTag = PR_ENTRYID;
    
    SizedENTRYID(16, TestEID);
    mapi_TEntryid mEID(16, (ENTRYID*)&TestEID);
    PropValue[0].Value.bin.cb = mEID.size;
    PropValue[0].Value.bin.lpb = mEID.ab->ab;
    
    TestRow.aRow[0].lpProps = PropValue;

    PropBase.FillData(&TestRow.aRow[0]);
    ASSERT(!PropBase.GetEntryID().isempty());
    ASSERT(PropBase.GetEntryID().size == 16);
  }

  static void GetPropValueTestCase() {
    TEST();
    CPropBase PropBase;
        
    SPropValue PropValue[3] = { NULL };
    
    PropValue[0].ulPropTag = PR_ENTRYID;    
    SizedENTRYID(16, TestEID);
    mapi_TEntryid mEID(16, (ENTRYID*)&TestEID);
    PropValue[0].Value.bin.cb = mEID.size;
    PropValue[0].Value.bin.lpb = mEID.ab->ab;    
    
    PropValue[1].ulPropTag = PR_DISPLAY_NAME;
    PropValue[1].Value.lpszA = "MyName";
    
    PropValue[2].ulPropTag = PR_COUNTRY;
    PropValue[2].Value.lpszA = "Poland";

    SRow TestRow = {NULL, 3, PropValue};    
    PropBase.FillData(&TestRow);

    ASSERT(!PropBase.GetEntryID().isempty());
    ASSERT(PropBase.GetEntryID().size == 16);

    ASSERT(PropBase.GetStringPropValue(PR_DISPLAY_NAME) == "MyName");
    ASSERT(PropBase.GetStringPropValue(PR_COUNTRY) == "Poland");
  }
  static void AddPropValueTestCase() {
    TEST();
    CPropBase PropBase;
    PropBase.AddStringPropValue(PR_DISPLAY_NAME, "MyName");
    PropBase.AddStringPropValue(PR_COUNTRY, "Poland");
    ASSERT(PropBase.GetStringPropValue(PR_DISPLAY_NAME) == "MyName");
    ASSERT(PropBase.GetStringPropValue(PR_COUNTRY) == "Poland");
  }
  static void PropTagExistsTestCase() {
    TEST();
    CPropBase PropBase;
    PropBase.AddStringPropValue(PR_DISPLAY_NAME, "MyName");
    PropBase.AddStringPropValue(PR_COUNTRY, "Poland");
    ASSERT(PropBase.PropStringTagExists(PR_DISPLAY_NAME));
    ASSERT(PropBase.PropStringTagExists(PR_COUNTRY));
  }  
  static void FillDataTestCase() {
    TEST();
    CPropBase PropBase;
    ASSERT(PropBase.GetEntryID().isempty());
    
    // Create sample row set.
    SizedSRowSet(1, TestRow) = {1, { NULL, NULL, NULL } };
    TestRow.aRow[0].cValues = 1;

    SPropValue PropValue[1] = { NULL };
    PropValue[0].ulPropTag = PR_ENTRYID;
    
    SizedENTRYID(16, TestEID);
    mapi_TEntryid mEID(16, (ENTRYID*)&TestEID);
    PropValue[0].Value.bin.cb = mEID.size;
    PropValue[0].Value.bin.lpb = mEID.ab->ab;
    
    TestRow.aRow[0].lpProps = PropValue;

    PropBase.FillData(&TestRow.aRow[0]);
    ASSERT(!PropBase.GetEntryID().isempty());
    ASSERT(PropBase.GetEntryID().size == 16);
  }
};

class CDistListTestCase {
public:
  static void AddMemberTestCase() {
    TEST();
    
    SizedENTRYID(16, TestEID1);
    mapi_TEntryid mEID1(16, (ENTRYID*)&TestEID1);
    mEID1.ab->ab[0] = 0x12;

    SizedENTRYID(16, TestEID2);
    mapi_TEntryid mEID2(16, (ENTRYID*)&TestEID2);
    mEID1.ab->ab[0] = 0x13;

    CDistList DistList;
    DistList.AddMember(mEID1);
    DistList.AddMember(mEID2);

    ASSERT(DistList.MembersCount() == 2);
    CWABWrapper WabWrap;
    WabWrap.Connect();
    ASSERT(DistList.GetMember(0).isequal(WabWrap.GetAdrBook(), mEID1));
    ASSERT(DistList.GetMember(1).isequal(WabWrap.GetAdrBook(), mEID2));
  }
  static void GetMemberTestCase() {
    TEST();
     
    SizedENTRYID(16, TestEID1);
    mapi_TEntryid mEID1(16, (ENTRYID*)&TestEID1);
    mEID1.ab->ab[0] = 0x12;

    SizedENTRYID(16, TestEID2);
    mapi_TEntryid mEID2(16, (ENTRYID*)&TestEID2);
    mEID1.ab->ab[0] = 0x13;

    CDistList DistList;
    DistList.AddMember(mEID1);
    DistList.AddMember(mEID2);

    ASSERT(DistList.MembersCount() == 2);
    CWABWrapper WabWrap;
    WabWrap.Connect();
    ASSERT(DistList.GetMember(0).isequal(WabWrap.GetAdrBook(), mEID1));
    ASSERT(DistList.GetMember(1).isequal(WabWrap.GetAdrBook(), mEID2));
  }
  static void MembersCountTestCase() {
    TEST();
     
    SizedENTRYID(16, TestEID1);
    mapi_TEntryid mEID1(16, (ENTRYID*)&TestEID1);
    mEID1.ab->ab[0] = 0x12;

    SizedENTRYID(16, TestEID2);
    mapi_TEntryid mEID2(16, (ENTRYID*)&TestEID2);
    mEID1.ab->ab[0] = 0x13;

    CDistList DistList;
    DistList.AddMember(mEID1);
    DistList.AddMember(mEID2);

    ASSERT(DistList.MembersCount() == 2);
  }
};

class CMailUserTestCase {
public:
  static void FillDataTestCase() {
    TEST();
    CMailUser MailUser;
    ASSERT(MailUser.GetEntryID().isempty());
    
    // Create sample row set, assume this is the data that WAB is returning.
    SizedSRowSet(1, TestRow) = {1, { NULL, NULL, NULL }};
    TestRow.aRow[0].cValues = 2;

    SPropValue PropValue[2] = { NULL };
    
    PropValue[0].ulPropTag = PR_ENTRYID;    
    SizedENTRYID(16, TestEID);
    mapi_TEntryid mEID(16, (ENTRYID*)&TestEID);
    PropValue[0].Value.bin.cb = mEID.size;
    PropValue[0].Value.bin.lpb = mEID.ab->ab;

    LPTSTR emails[2] = { TEXT("email1@www.com"), TEXT("email2@www.com") };
    PropValue[1].ulPropTag = PR_CONTACT_EMAIL_ADDRESSES;
    PropValue[1].Value.MVszA.lppszA = emails;
    PropValue[1].Value.MVszA.cValues = 2;
    
    TestRow.aRow[0].lpProps = PropValue;

    // Now fill MailUser object with data from WAB.
    MailUser.FillData(&TestRow.aRow[0]);

    // Verify that MailUser object was initialized properly.
    ASSERT(!MailUser.GetEntryID().isempty());
    ASSERT(MailUser.GetEntryID().size == 16);
    
    _STL::vector<_STL::string> aEmails;
    HRESULT hr = MailUser.GetStringPropMultiValue(PR_CONTACT_EMAIL_ADDRESSES, aEmails);
    ASSERT(hr == S_OK);
    ASSERT(aEmails.size() == 2);
    ASSERT(aEmails[0] == emails[0]);
    ASSERT(aEmails[1] == emails[1]);
  }
};

class CWABWrapperTestCase {
public:
  static void ConnectTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);
  }
  static void GetAdrBookTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);
    ASSERT(WabWrap.GetAdrBook() != NULL);
  }
  static void DisconnectTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);
    hr = WabWrap.Disconnect();
    ASSERT(hr == S_OK);
  }
  static void IsConnectedTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);
    ASSERT(WabWrap.IsConnected());
    hr = WabWrap.Disconnect();
    ASSERT(hr == S_OK);
  }
  static void PropertiesTestCase() {
    TEST();
    CWABWrapper WabWrap;
    WabWrap.Properties().Add(PR_DISPLAY_NAME);    
    ASSERT(WabWrap.Properties()[0] == PR_DISPLAY_NAME);
  }
  static void GetFoldersTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);
    _STL::vector<_STL::string> aFolders;
    hr = WabWrap.GetFolders(aFolders);
    ASSERT(hr == S_OK);
    ASSERT(aFolders.size() >= 1);
  }
  static void SetFoldersTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);
    _STL::vector<_STL::string> aFolders;
    hr = WabWrap.GetFolders(aFolders);
    ASSERT(hr == S_OK);
    ASSERT(aFolders.size() >= 1);
    hr = WabWrap.SetFolder(aFolders[0]);
    ASSERT(hr == S_OK);
  }
  static void SetPABFolderTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);    
    hr = WabWrap.SetPABFolder();
    ASSERT(hr == S_OK);    
  }
  static void CreateFolderTestCase() {
    TEST();
    CWABWrapper WabWrap;
    HRESULT hr = WabWrap.Connect();
    ASSERT(hr == S_OK);    
    hr = WabWrap.SetPABFolder();
    ASSERT(hr == S_OK);    

    WabWrap.CreateFolder("Test21");
  }
  static void GetMAPIErrorTestCase() {
    TEST();    
    ASSERT(0 != stricmp(CWABWrapper::GetMAPIError(S_OK).c_str(), "Unknown"));
    ASSERT(0 != stricmp(CWABWrapper::GetMAPIError(MAPI_E_COMPUTED).c_str(), "Unknown"));
    ASSERT(0 != stricmp(CWABWrapper::GetMAPIError(MAPI_E_BAD_CHARWIDTH).c_str(), "Unknown"));
    ASSERT(0 != stricmp(CWABWrapper::GetMAPIError(MAPI_E_AMBIGUOUS_RECIP).c_str(), "Unknown"));
  }
  static void SetFilterTestCase() {
    TEST();
    ASSERT(false);
  }
  static void SetFilterWithPropsTestCase() {
    TEST();
    ASSERT(false);
  }
  static void GetFolderItemCountTestCase() {
    TEST();
    ASSERT(false);
  }
  static void GetItemTypeByIndexTestCase() {
    TEST();
    ASSERT(false);
  }
  static void GetItemTypeByEntryIdTestCase() {
    TEST();
    ASSERT(false);
  }
  static void GetDistListTestCase()  {
    TEST();
    ASSERT(false);
  }
  static void UpdateDistListTestCase() {
    TEST();
    ASSERT(false);
  }
  static void AddDistListTestCase() {
    TEST();
    ASSERT(false);
  }
  static void RemoveDistListTestCase() {
    TEST();
    ASSERT(false);
  }
  static void GetMailUserByIndexTestCase() {
    TEST();
    ASSERT(false);
  }
  static void GetMailUserByEntryTestCase() {
    TEST();
    ASSERT(false);
  }
  static void UpdateMailUserTestCase() {
    TEST();
    ASSERT(false);
  }
  static void AddMailUserTestCase() {
    TEST();
    ASSERT(false);
  }
  static void RemoveMailUserTestCase() {
    TEST();
    ASSERT(false);
  }
};

int main(int argc, char* argv[])
{
/*
  // CDynamicPropertyTagArray
  CDynamicPropertyTagArrayTestCase::AddTestCase();
  CDynamicPropertyTagArrayTestCase::CountTestCase();
  CDynamicPropertyTagArrayTestCase::ClearTestCase();
  CDynamicPropertyTagArrayTestCase::LPSPropTagArrayOpTestCase();
  CDynamicPropertyTagArrayTestCase::IndexOpTestCase();

  // CPropBase
  CPropBaseTestCase::GetEntryIDTestCase();
  CPropBaseTestCase::GetPropValueTestCase();
  CPropBaseTestCase::AddPropValueTestCase();
  CPropBaseTestCase::PropTagExistsTestCase();
  CPropBaseTestCase::FillDataTestCase();

  // CDistList 
  CDistListTestCase::AddMemberTestCase();
  CDistListTestCase::GetMemberTestCase();
  CDistListTestCase::MembersCountTestCase();

  // CMailUser
  CMailUserTestCase::FillDataTestCase();

  // CWABWrapper
  CWABWrapperTestCase::ConnectTestCase();
  CWABWrapperTestCase::GetAdrBookTestCase();
  CWABWrapperTestCase::DisconnectTestCase();
  CWABWrapperTestCase::IsConnectedTestCase();
  CWABWrapperTestCase::PropertiesTestCase();
  CWABWrapperTestCase::GetFoldersTestCase();
  CWABWrapperTestCase::SetFoldersTestCase();
  CWABWrapperTestCase::SetPABFolderTestCase();
*/
  CWABWrapperTestCase::CreateFolderTestCase();
  CWABWrapperTestCase::GetMAPIErrorTestCase();

  CWABWrapperTestCase::SetFilterTestCase();
  CWABWrapperTestCase::SetFilterWithPropsTestCase();  

  CWABWrapperTestCase::GetFolderItemCountTestCase();

  CWABWrapperTestCase::GetItemTypeByIndexTestCase();
  CWABWrapperTestCase::GetItemTypeByEntryIdTestCase();

  CWABWrapperTestCase::GetDistListTestCase();  
  CWABWrapperTestCase::UpdateDistListTestCase();
  CWABWrapperTestCase::AddDistListTestCase();
  CWABWrapperTestCase::RemoveDistListTestCase();

  CWABWrapperTestCase::GetMailUserByIndexTestCase();
  CWABWrapperTestCase::GetMailUserByEntryTestCase();
  CWABWrapperTestCase::UpdateMailUserTestCase();
  CWABWrapperTestCase::AddMailUserTestCase();
  CWABWrapperTestCase::RemoveMailUserTestCase(); 

	return 0;
}

