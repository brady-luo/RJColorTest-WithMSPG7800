#include "pch.h"


#include "framework.h"
#include "RJColorTestMSPG.h"
#include "RJColorTestMSPGDlg.h"
#include "afxdialogex.h"


#include "CaEvent.h"

// CA-SDK


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCaEvent

IMPLEMENT_DYNCREATE(CCaEvent, CCmdTarget)

CCaEvent::CCaEvent()
{
	EnableAutomation();    //自动化函数
}

CCaEvent::~CCaEvent()
{
}


void CCaEvent::OnFinalRelease()
{

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CCaEvent, CCmdTarget)
	//{{AFX_MSG_MAP(CCaEvent)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CCaEvent, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CCaEvent)
	DISP_FUNCTION(CCaEvent, "ExeCalZero", ExeCalZero, VT_EMPTY, VTS_NONE)
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()


// CA-SDK
// {D497071B-3F9F-41C2-9E04-1E573571BF0D}
// static const IID IID_ICaEvent =
// { 0xd497071b, 0x3f9f, 0x41c2, { 0x9e, 0x4, 0x1e, 0x57, 0x35, 0x71, 0xbf, 0xd } };
// {f7663750-5900-45eb-905f-78c5d5378481}
const IID IID_ICaEvent =
{0xf7663750,0x5900,0x45eb,{0x90,0x5f,0x78,0xc5,0xd5,0x37,0x84,0x81}};

BEGIN_INTERFACE_MAP(CCaEvent, CCmdTarget)
	INTERFACE_PART(CCaEvent, IID_ICaEvent, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCaEvent Message handler

void CCaEvent::ExeCalZero() 
{
	// CA-SDK
	CWinApp* papp = AfxGetApp();
	CRJColorTestMSPGDlg* pdlg = reinterpret_cast<CRJColorTestMSPGDlg*>(papp -> m_pMainWnd);

	//CButton* pb;
	//pb = (CButton *)(pdlg ->GetDlgItem(IDC_BTN_BEGIN));  //开始调节按钮
	//pb->EnableWindow(FALSE);  //开始调节按钮变为灰

}

//Copyright (c) 2002-2010 KONICA MINOLTA SENSING, INC. 