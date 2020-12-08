
// RJColorTestMSPGDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "RJColorTestMSPG.h"
#include "RJColorTestMSPGDlg.h"
#include "afxdialogex.h"

#include <iostream>
#include <fstream>
#include <string>
#include<time.h>  

//EnumsCOMS
#include <winioctl.h>
#include<Setupapi.h>
#pragma comment(lib, "Setupapi.lib")

// CA-SDK
#include "Const.h"       
#include "CaEvent.h"     
#import "C:\Program Files (x86)\KONICAMINOLTA\CA-SDK\SDK\CA200Srvr.dll" no_namespace implementation_only   //编译后会自动生成CA200Srvr.tlh

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// The following define is from ntddser.h in the DDK. 用于枚举出串口在设备管理器显示的名称
#ifndef GUID_CLASS_COMPORT
DEFINE_GUID(GUID_CLASS_COMPORT, 0x86e0d1e0L, 0x8089, 0x11d0, 0x9c, 0xe4, \
	0x08, 0x00, 0x3e, 0x30, 0x1f, 0x73);
#endif


using namespace std;

//全局变量
string str_Gray, str_fx, str_fy, str_Lv, str_CCT;
ofstream outUserTestFile;
int step7800; //7800播放时间间隔
HANDLE g_hWaitContrast, g_hWaitColor, g_hWaitGamma;

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CRJColorTestMSPGDlg 对话框


CRJColorTestMSPGDlg::CRJColorTestMSPGDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_RJCOLORTEST_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);
}

void CRJColorTestMSPGDlg::DoDataExchange(CDataExchange* pDX)
{
	//CDialogEx::DoDataExchange(pDX);
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_OpenCA, m_isOpencaOK);
	DDX_Control(pDX, IDC_LIST_TestResult, m_isTestOver);
	DDX_Control(pDX, IDC_COMBO_Com, m_ComNumber);
	DDX_Control(pDX, IDC_EDIT_ShowC, m_EditShowC);
	DDX_Control(pDX, IDC_COMBO_7800Step, vCombo_7800Step);
}

BEGIN_MESSAGE_MAP(CRJColorTestMSPGDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_GammaTest, &CRJColorTestMSPGDlg::OnBnClickedBtnGammatest)
	ON_BN_CLICKED(IDC_BTN_ColorTest, &CRJColorTestMSPGDlg::OnBnClickedBtnColortest)
	ON_BN_CLICKED(IDC_BTN_ContrastTest, &CRJColorTestMSPGDlg::OnBnClickedBtnContrasttest)
	ON_BN_CLICKED(IDC_BTN_AllTest, &CRJColorTestMSPGDlg::OnBnClickedBtnAlltest)
	ON_BN_CLICKED(IDC_BTN_ConnectCA, &CRJColorTestMSPGDlg::OnBnClickedBtnConnectca)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_OpenCA, &CRJColorTestMSPGDlg::OnNMCustomdrawListOpenca)
	ON_BN_CLICKED(IDOK, &CRJColorTestMSPGDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CRJColorTestMSPGDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BTN_UserTest, &CRJColorTestMSPGDlg::OnBnClickedBtnUsertest)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_TestResult, &CRJColorTestMSPGDlg::OnNMCustomdrawListTestresult)
	ON_BN_CLICKED(IDC_CHECK_IsSaveData, &CRJColorTestMSPGDlg::OnBnClickedCheckIssavedata)
	ON_BN_CLICKED(IDC_BTN_FindCOM, &CRJColorTestMSPGDlg::OnBnClickedBtnFindcom)
	ON_BN_CLICKED(IDC_BTN_OpenCOM, &CRJColorTestMSPGDlg::OnBnClickedBtnOpencom)
	ON_BN_CLICKED(IDC_BTN_Connect7800, &CRJColorTestMSPGDlg::OnBnClickedBtnConnect7800)
	ON_BN_CLICKED(IDC_BTN_Close7800, &CRJColorTestMSPGDlg::OnBnClickedBtnClose7800)
	ON_BN_CLICKED(IDC_BTN_SendTimingC, &CRJColorTestMSPGDlg::OnBnClickedBtnSendtimingc)
	ON_BN_CLICKED(IDC_BTN_SendPatternC, &CRJColorTestMSPGDlg::OnBnClickedBtnSendpatternc)
	ON_BN_CLICKED(IDC_BTN_SendUserC, &CRJColorTestMSPGDlg::OnBnClickedBtnSenduserc)
	ON_CBN_SELCHANGE(IDC_COMBO_7800Step, &CRJColorTestMSPGDlg::OnCbnSelchangeCombo7800step)
END_MESSAGE_MAP()


// CRJColorTestMSPGDlg 消息处理程序

BOOL CRJColorTestMSPGDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	//初始化设置listControl
	CFont m_Font;      // 字体变量 m_Font 的声明。只能声明成局部变量，否则assert failed
	m_Font.CreatePointFont(150, _T("新宋体"));  // 设置字体的大小和格式
	m_isOpencaOK.SetFont(&m_Font, true);
	m_isOpencaOK.InsertColumn(0, _T("第一列"), LVCFMT_CENTER, 125);
	m_isOpencaOK.InsertItem(0, _T("请连接CA310"), LVCFMT_CENTER);
	m_isOpencaOK.SetItemData(0, 3);

	m_isTestOver.SetFont(&m_Font, true);
	m_isTestOver.InsertColumn(0, _T("第一列"), LVCFMT_CENTER, 140);
	m_isTestOver.InsertItem(0, _T("  未开始测试  "), LVCFMT_CENTER);
	m_isTestOver.SetItemData(0, 3);

	//m_pictureName.SetWindowText(_T("请填写画面名称"));
	SetDlgItemText(IDC_EDIT_PictureName, _T("请填写画面名称"));

	//ChannelNO设置初始值，默认为0 (Konica Minolta calibration)
	SetDlgItemText(IDC_ChannelNO, _T("0"));

	//初始化控件不可用
	HWND hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_GammaTest);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ColorTest);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ContrastTest);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_AllTest);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_UserTest);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ConnectCA);
	::EnableWindow(hBtn, FALSE);

	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_Connect7800);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_Close7800);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendTimingC);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendPatternC);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendUserC);
	::EnableWindow(hBtn, FALSE);

	//初始化时间步长
	vCombo_7800Step.SetCurSel(1);
	CString curStr_7800Step;
	vCombo_7800Step.GetLBText(1, curStr_7800Step);
	step7800 = _ttoi(curStr_7800Step);

	//获取可用的串口
	OnBnClickedBtnFindcom();
	//初始化COM属性下拉框
	((CComboBox*)GetDlgItem(IDC_COMBO_BaudRate))->SetCurSel(4);
	((CComboBox*)GetDlgItem(IDC_COMBO_DataBit))->SetCurSel(1);
	((CComboBox*)GetDlgItem(IDC_COMBO_StopBit))->SetCurSel(0);
	((CComboBox*)GetDlgItem(IDC_COMBO_CheckBit))->SetCurSel(0);

	//默认只测试W Gamma，勾选
	((CButton*)GetDlgItem(IDC_CHECK_isTestW))->SetCheck(1);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CRJColorTestMSPGDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CRJColorTestMSPGDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CRJColorTestMSPGDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



DWORD CRJColorTestMSPGDlg::ThreadFunc_Gamma(LPVOID lpParam)
{
	DWORD dWait = WaitForSingleObject(g_hWaitGamma, INFINITE);//无限等待
	if (WAIT_OBJECT_0 == dWait)
	{
		CRJColorTestMSPGDlg* PDLG = (CRJColorTestMSPGDlg*)lpParam;
		//显示正在状态
		PDLG->m_EditShowC.SetWindowTextW(_T(""));//清空命令显示框，避免超出最大行数限制
		PDLG->m_EditShowC.SetSel(-1);
		PDLG->m_EditShowC.ReplaceSel(_T("---GammaTest进行中!----\r\n"));
		PDLG->m_isTestOver.SetItemText(0, 0, _T("GammaTest进行中..."));
		PDLG->m_isTestOver.SetItemData(0, 2);  //设置红底
		//获取系统时间，作为文件名称	
		CString timer = PDLG->timeChangeFormat();
		string filename = "TestResult\\" + CStringA(timer) + "-GammaTest.csv";
		//打开文件
		ofstream outGammaFile;
		outGammaFile.open(filename, ios::out);
		outGammaFile << "Gray" << ',' << "fx" << ',' << "fy" << ',' << "Lv" << ',' << "CCT" << endl;

		//7800打开 Window pattern 31
		BYTE commandPattern[] = { 0x02,0x08,0x30,0x33,0x31,0x03,'\0' };
		int commandPatternLen = sizeof(commandPattern) - 1;
		PDLG->sendCommand(commandPattern, commandPatternLen);
		Sleep(step7800);
		Sleep(step7800);

		//7800调整pattern的Window大小为100% @H @V
		BYTE commandBoxSize[] = { 0x02,0x11,0x31,0x30,0x30,0x31,0x30,0x30,0x03,'\0' };
		int commandBoxSizeLen = sizeof(commandBoxSize) - 1;
		PDLG->sendCommand(commandBoxSize, commandBoxSizeLen);
		Sleep(step7800);
		Sleep(step7800);

		//测试Gamma
		char temp1, temp2, temp3; //用于表示个位十位百位

		//先测W Gamma
		int isTestWGamma = ((CButton*)(PDLG->GetDlgItem(IDC_CHECK_isTestW)))->GetCheck();
		if (isTestWGamma == 1) {
			BYTE cGray[] = { 0x02,0x10,0x30,0x30,0x30,0x30,
										0x30,0x30,0x30,0x30,
										0x30,0x30,0x30,0x30,0x03,'\0' };
			int cGrayLen = sizeof(cGray) - 1;
			PDLG->sendCommand(cGray, cGrayLen); //先测第一个画面
			Sleep(step7800);
			Sleep(step7800);
			PDLG->CA_Measure_SxSyLv();
			str_Gray = "W";
			outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

			for (int i = 3; i <= 255;) {
				if (i < 10) {
					temp1 = (char)(i + 48);
					cGray[3] = cGray[7] = cGray[11] = '0';
					cGray[4] = cGray[8] = cGray[12] = '0';
					cGray[5] = cGray[9] = cGray[13] = temp1;
				}
				else if (i >= 10 && i < 100) {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)(i / 10 + 48);
					cGray[3] = cGray[7] = cGray[11] = '0';
					cGray[4] = cGray[8] = cGray[12] = temp2;
					cGray[5] = cGray[9] = cGray[13] = temp1;
				}
				else {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)((i / 10) % 10 + 48);
					temp3 = (char)(i / 100 + 48);
					cGray[3] = cGray[7] = cGray[11] = temp3;
					cGray[4] = cGray[8] = cGray[12] = temp2;
					cGray[5] = cGray[9] = cGray[13] = temp1;
				}

				PDLG->sendCommand(cGray, cGrayLen);
				Sleep(step7800);
				PDLG->CA_Measure_SxSyLv();
				str_Gray = to_string(i);
				outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

				i += 4;
			}
		}
		

		//测B Gamma
		int isTestBGamma = ((CButton*)(PDLG->GetDlgItem(IDC_CHECK_isTestB)))->GetCheck();
		if (isTestBGamma == 1) {
			BYTE cBGray[] = { 0x02,0x10,0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,0x03,'\0' };
			int cBGrayLen = sizeof(cBGray) - 1;
			PDLG->sendCommand(cBGray, cBGrayLen);
			Sleep(step7800);
			Sleep(step7800);
			PDLG->CA_Measure_SxSyLv();
			str_Gray = "B";
			outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

			for (int i = 3; i <= 255;) {
				if (i < 10) {
					temp1 = (char)(i + 48);
					cBGray[11] = '0';
					cBGray[12] = '0';
					cBGray[13] = temp1;
				}
				else if (i >= 10 && i < 100) {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)(i / 10 + 48);
					cBGray[11] = '0';
					cBGray[12] = temp2;
					cBGray[13] = temp1;
				}
				else {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)((i / 10) % 10 + 48);
					temp3 = (char)(i / 100 + 48);
					cBGray[11] = temp3;
					cBGray[12] = temp2;
					cBGray[13] = temp1;
				}

				PDLG->sendCommand(cBGray, cBGrayLen);
				Sleep(step7800);
				PDLG->CA_Measure_SxSyLv();
				str_Gray = to_string(i);
				outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

				i += 4;
			}
		}
		


		//测G Gamma
		int isTestGGamma = ((CButton*)(PDLG->GetDlgItem(IDC_CHECK_isTestG)))->GetCheck();
		if (isTestGGamma == 1) {
			BYTE cGGray[] = { 0x02,0x10,0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,0x03,'\0' };
			int cGGrayLen = sizeof(cGGray) - 1;
			PDLG->sendCommand(cGGray, cGGrayLen);
			Sleep(step7800);
			Sleep(step7800);
			PDLG->CA_Measure_SxSyLv();
			str_Gray = "G";
			outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

			for (int i = 3; i <= 255;) {
				if (i < 10) {
					temp1 = (char)(i + 48);
					cGGray[7] = '0';
					cGGray[8] = '0';
					cGGray[9] = temp1;
				}
				else if (i >= 10 && i < 100) {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)(i / 10 + 48);
					cGGray[7] = '0';
					cGGray[8] = temp2;
					cGGray[9] = temp1;
				}
				else {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)((i / 10) % 10 + 48);
					temp3 = (char)(i / 100 + 48);
					cGGray[7] = temp3;
					cGGray[8] = temp2;
					cGGray[9] = temp1;
				}

				PDLG->sendCommand(cGGray, cGGrayLen);
				Sleep(step7800);
				PDLG->CA_Measure_SxSyLv();
				str_Gray = to_string(i);
				outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

				i += 4;
			}
		}
		


		//测R Gamma
		int isTestRGamma = ((CButton*)(PDLG->GetDlgItem(IDC_CHECK_isTestR)))->GetCheck();
		if (isTestRGamma == 1) {
			BYTE cRGray[] = { 0x02,0x10,0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,0x03,'\0' };
			int cRGrayLen = sizeof(cRGray) - 1;
			PDLG->sendCommand(cRGray, cRGrayLen);
			Sleep(step7800);
			Sleep(step7800);
			PDLG->CA_Measure_SxSyLv();
			str_Gray = "R";
			outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

			for (int i = 3; i <= 255;) {
				if (i < 10) {
					temp1 = (char)(i + 48);
					cRGray[3] = '0';
					cRGray[4] = '0';
					cRGray[5] = temp1;
				}
				else if (i >= 10 && i < 100) {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)(i / 10 + 48);
					cRGray[3] = '0';
					cRGray[4] = temp2;
					cRGray[5] = temp1;
				}
				else {
					temp1 = (char)(i % 10 + 48);
					temp2 = (char)((i / 10) % 10 + 48);
					temp3 = (char)(i / 100 + 48);
					cRGray[3] = temp3;
					cRGray[4] = temp2;
					cRGray[5] = temp1;
				}

				PDLG->sendCommand(cRGray, cRGrayLen);
				Sleep(step7800);
				PDLG->CA_Measure_SxSyLv();
				str_Gray = to_string(i);
				outGammaFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

				i += 4;
			}
		}

		outGammaFile.close();
		//显示完成状态
		PDLG->m_EditShowC.SetSel(-1);
		PDLG->m_EditShowC.ReplaceSel(_T("---GammaTest完毕!----\r\n"));
		PDLG->m_isTestOver.SetItemText(0, 0, _T("GammaTest完成"));
		PDLG->m_isTestOver.SetItemData(0, 1);  //设置绿底
	}
	return 0;
}

void CRJColorTestMSPGDlg::OnBnClickedBtnGammatest()
{
	g_hWaitGamma = CreateEvent(NULL, TRUE, TRUE, NULL);//创建一个初始为有信号的事件量
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadFunc_Gamma, this, 0, 0);
}


DWORD CRJColorTestMSPGDlg::ThreadFunc_Color(LPVOID lpParam)
{
	DWORD dWait = WaitForSingleObject(g_hWaitColor, INFINITE);//无限等待
	if (WAIT_OBJECT_0 == dWait)
	{

		CRJColorTestMSPGDlg* PDLG = (CRJColorTestMSPGDlg*)lpParam;

		//显示正在状态
		PDLG->m_EditShowC.SetSel(-1);
		PDLG->m_EditShowC.ReplaceSel(_T("---ColorTest进行中!----\r\n"));
		PDLG->m_isTestOver.SetItemText(0, 0, _T("ColorTest进行中..."));
		PDLG->m_isTestOver.SetItemData(0, 2);  //设置红底
		//获取系统时间，作为文件名称
		CString timer = PDLG->timeChangeFormat();
		string filename = "TestResult\\" + CStringA(timer) + "-ColorTest.csv";
		//打开文件
		ofstream outColorFile;
		outColorFile.open(filename, ios::out);
		outColorFile << "Picture" << ',' << "fx" << ',' << "fy" << ',' << "Lv" << ',' << "CCT" << endl;

		//7800打开 Window pattern 31
		BYTE commandPattern[] = { 0x02,0x08,0x31,0x33,0x36,0x03,'\0' };
		int commandPatternLen = sizeof(commandPattern) - 1;
		PDLG->sendCommand(commandPattern, commandPatternLen);
		Sleep(step7800);
		Sleep(step7800);

		//7800调整pattern的Window大小为100% @H @V
		BYTE commandBoxSize[] = { 0x02,0x11,0x31,0x30,0x30,0x31,0x30,0x30,0x03,'\0' };
		int commandBoxSizeLen = sizeof(commandBoxSize) - 1;
		PDLG->sendCommand(commandBoxSize, commandBoxSizeLen);
		Sleep(step7800);
		Sleep(step7800);

		//先打一个黑色画面
		BYTE cGray[] = { 0x02,0x10,0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,
									0x30,0x30,0x30,0x30,0x03,'\0' };
		int cGrayLen = sizeof(cGray) - 1;
		PDLG->sendCommand(cGray, cGrayLen);
		Sleep(step7800);
		Sleep(step7800);

		CString strline;
		CStdioFile file;
		BOOL flag = file.Open(_T("colorValue1.txt"), CFile::modeRead);
		if (flag == FALSE) {
			AfxMessageBox(_T("颜色值文件打开失败！"));
			return 0;
		}
		while (file.ReadString(strline)) {
			if (strline.GetLength() != 13) {
				AfxMessageBox(_T("颜色值文件内容不正确！"));
				return 0;
			}
			cGray[3] = (BYTE)strline[2];//R
			cGray[4] = (BYTE)strline[3];
			cGray[5] = (BYTE)strline[4];

			cGray[7] = (BYTE)strline[6];//G
			cGray[8] = (BYTE)strline[7];
			cGray[9] = (BYTE)strline[8];

			cGray[11] = (BYTE)strline[10];//B
			cGray[12] = (BYTE)strline[11];
			cGray[13] = (BYTE)strline[12];

			PDLG->sendCommand(cGray, cGrayLen);
			Sleep(step7800);
			PDLG->CA_Measure_SxSyLv();
			str_Gray = (char)strline[0];
			outColorFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

		}

		outColorFile.close();
		//显示完成状态
		PDLG->m_EditShowC.SetSel(-1);
		PDLG->m_EditShowC.ReplaceSel(_T("---ColorTest完成!----\r\n"));
		PDLG->m_isTestOver.SetItemText(0, 0, _T("ColorTest完成"));
		PDLG->m_isTestOver.SetItemData(0, 1);  //设置绿底 

		SetEvent(g_hWaitGamma);//转为有信号状态，其他线程的WaitForSingleObject会返回WAIT_OBJECT_0
	}
	return 0;
}

void CRJColorTestMSPGDlg::OnBnClickedBtnColortest()
{
	g_hWaitColor = CreateEvent(NULL, TRUE, TRUE, NULL);//创建一个初始为有信号的事件量
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadFunc_Color, this, 0, 0);
}


DWORD CRJColorTestMSPGDlg::ThreadFunc_Contrast(LPVOID lpParam)
{
	SetEvent(g_hWaitContrast);
	DWORD dWait = WaitForSingleObject(g_hWaitContrast, INFINITE);//无限等待
	if (WAIT_OBJECT_0 == dWait)
	{
		CRJColorTestMSPGDlg* PDLG = (CRJColorTestMSPGDlg*)lpParam;

		//显示正在状态
		PDLG->m_EditShowC.SetSel(-1);
		PDLG->m_EditShowC.ReplaceSel(_T("---ContrastTest进行中!----\r\n"));
		PDLG->m_isTestOver.SetItemText(0, 0, _T("ContrastTest进行中..."));
		PDLG->m_isTestOver.SetItemData(0, 2);  //设置红底
		//获取系统时间，作为文件名称	
		CString timer = PDLG->timeChangeFormat();
		string filename = "TestResult\\" + CStringA(timer) + "-ContrastTest.csv";
		//打开文件
		ofstream outContrastFile;
		outContrastFile.open(filename, ios::out);
		outContrastFile << "Gray" << ',' << "fx" << ',' << "fy" << ',' << "Lv" << ',' << "CCT" << endl;


		//7800打开 W pattern 004
		BYTE commandPattern[] = { 0x02,0x08,0x30,0x30,0x34,0x03,'\0' };
		int commandPatternLen = sizeof(commandPattern) - 1;
		PDLG->sendCommand(commandPattern, commandPatternLen);
		Sleep(step7800);
		Sleep(step7800);
		PDLG->CA_Measure_SxSyLv();
		str_Gray = "White";
		outContrastFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;

		//7800打开 Black pattern 002
		commandPattern[4] = { 0x32 };
		PDLG->sendCommand(commandPattern, commandPatternLen);
		Sleep(step7800);
		Sleep(step7800);
		PDLG->CA_Measure_SxSyLv();
		str_Gray = "Black";
		outContrastFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;


		outContrastFile.close();
		//显示完成状态
		PDLG->m_EditShowC.SetSel(-1);
		PDLG->m_EditShowC.ReplaceSel(_T("---ContrastTest完成!----\r\n"));
		PDLG->m_isTestOver.SetItemText(0, 0, _T("ContrastTest完成"));
		PDLG->m_isTestOver.SetItemData(0, 1);  //设置绿底

		SetEvent(g_hWaitColor);//转为有信号状态，其他线程的WaitForSingleObject会返回WAIT_OBJECT_0
	}
	return 0;
}

void CRJColorTestMSPGDlg::OnBnClickedBtnContrasttest()
{
	g_hWaitContrast = CreateEvent(NULL, TRUE, TRUE, NULL);//创建一个初始为有信号的事件量
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadFunc_Contrast, this, 0, 0);
}


void CRJColorTestMSPGDlg::OnBnClickedBtnAlltest()
{
	g_hWaitContrast = CreateEvent(NULL, TRUE, FALSE, NULL);//创建一个手动复位、初始为无信号的事件
	g_hWaitColor = CreateEvent(NULL, TRUE, FALSE, NULL);//创建一个手动复位、初始为无信号的事件
	g_hWaitGamma = CreateEvent(NULL, TRUE, FALSE, NULL);//创建一个手动复位、初始为无信号的事件
	DWORD ThreadID_Contrast, ThreadID_Color, ThreadID_Gamma;
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadFunc_Contrast, this, 0, &ThreadID_Contrast);
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadFunc_Color, this, 0, &ThreadID_Color);
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadFunc_Gamma, this, 0, &ThreadID_Gamma);
}


void CRJColorTestMSPGDlg::OnBnClickedBtnUsertest()
{
	CString CStr_pictureName;
	GetDlgItem(IDC_EDIT_PictureName)->GetWindowText(CStr_pictureName);
	str_Gray = CStringA(CStr_pictureName);
	CA_Measure_SxSyLv();

	if (((CButton*)GetDlgItem(IDC_CHECK_IsSaveData))->GetCheck())
	{
		outUserTestFile << str_Gray << ',' << str_fx << ',' << str_fy << ',' << str_Lv << ',' << str_CCT << endl;
	}

	Sleep(800);//等待800ms才更改编辑框文字
	SetDlgItemText(IDC_EDIT_PictureName, _T("请填写画面名称"));
}


void CRJColorTestMSPGDlg::OnBnClickedCheckIssavedata()
{
	int state = ((CButton*)GetDlgItem(IDC_CHECK_IsSaveData))->GetCheck();
	if (state == 1)
	{
		//获取系统时间，作为文件名称	
		CString timer = timeChangeFormat();
		string filename = "TestResult\\" + CStringA(timer) + "-UserTest.csv";
		//打开自定义测试的保存文件	
		outUserTestFile.open(filename, ios::app);
		outUserTestFile << "Picture" << ',' << "fx" << ',' << "fy" << ',' << "Lv" << ',' << "CCT" << endl;
	}
	else if (state == 0)
	{
		outUserTestFile.close();
	}
}


void CRJColorTestMSPGDlg::OnBnClickedBtnConnectca()
{
	// TODO: 在此添加控件通知处理程序代码
	// CA-SDK  创建和初始化 CA-SDK对象。使用SetConfiguration method
	long lcan = 1;
	_bstr_t strcnfig(_T("1"));
	long lprt = PORT_USB;
	long lbr = 38400;        //读取波特率设置
	_bstr_t strprbid(_T("P1"));
	_variant_t vprbid(_T("P1"));

	try {

		m_pCa200Obj = ICa200Ptr(__uuidof(Ca200));   // 1、创建CA200对象
		m_pCa200Obj->SetConfiguration(lcan, strcnfig, lprt, lbr);

		/*2、使用SetConfiguration()方法，通过以下设置将程序设置为使用1个CA单元
			CA unit CA number = 1
			Probe: 1 connected, probe number = 1
			USB connection
		*/
	}
	catch (_com_error e) {
		AfxMessageBox(_T("打开CA310失败!"));
		//return TRUE;
	}

	//在CA- sdk对象级别上执行用于控制CA单元的对象的设置
	m_pCasObj = m_pCa200Obj->Cas;      //从Ca200对象获取Cas集合
	m_pCaObj = m_pCasObj->ItemOfNumber[lcan];  //
	m_pOutputProbesObj = m_pCaObj->OutputProbes; //
	m_pOutputProbesObj->RemoveAll();          //
	m_pOutputProbesObj->Add(strprbid);     //
	m_pProbeObj = m_pOutputProbesObj->Item[vprbid]; //
	m_pMemoryObj = m_pCaObj->Memory;        //
	//各种对象用于执行以下CA单元初始化和内存通道设置
	m_pCaObj->SyncMode = SYNC_NTSC;   //将同步模式设置为NTSC
	m_pCaObj->AveragingMode = AVRG_FAST;  //设置快/慢模式为快
	m_pCaObj->SetAnalogRange(2.5, 2.5);  //设置模拟显示范围为2.5%，2.5%
	m_pCaObj->DisplayMode = DSP_LXY;     //设置显示模式
	m_pCaObj->DisplayDigits = DIGT_4;    //设置显示数字的数目为4

	//Set memory channel
	CString CStr_ChannelNO;
	GetDlgItem(IDC_ChannelNO)->GetWindowText(CStr_ChannelNO);
	m_pMemoryObj->ChannelNO = _ttol(CStr_ChannelNO);


	// CA-SDK 0 Cal校准
	CFont m_Font;      // 字体变量 m_Font 的声明。只能声明成局部变量，否则assert failed
	m_Font.CreatePointFont(100, _T("新宋体"));  // 设置字体的大小和格式
	//vResult_OpenCOM.SetFont(&m_Font, true);   // 设置OpenCOM列表框中的字体的格式
	m_isOpencaOK.SetFont(&m_Font, true);

	try {
		m_pCaObj->CalZero();
		m_isOpencaOK.SetItemText(0, 0, _T("CA310校准成功！"));
		m_isOpencaOK.SetItemData(0, 1);  //设置绿底

	}
	catch (_com_error e) {

		m_isOpencaOK.SetItemText(0, 0, _T("CA310校准失败！"));
		m_isOpencaOK.SetItemData(0, 2);  //设置红底

		AfxMessageBox(_T("CA310校准失败!"));
		return;
	}
	//CA连接OK
	CAIsOK = true;
	//测试控件设置为可用
	HWND hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_GammaTest);
	::EnableWindow(hBtn, TRUE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ColorTest);
	::EnableWindow(hBtn, TRUE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ContrastTest);
	::EnableWindow(hBtn, TRUE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_AllTest);
	::EnableWindow(hBtn, TRUE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_UserTest);
	::EnableWindow(hBtn, TRUE);
	//连接CA310控件 设置为不可操作
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ConnectCA);
	::EnableWindow(hBtn, FALSE);
}


void CRJColorTestMSPGDlg::OnNMCustomdrawListOpenca(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMTVCUSTOMDRAW pNMCD = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);
	NMCUSTOMDRAW nmCustomDraw = pNMCD->nmcd;
	switch (nmCustomDraw.dwDrawStage)
	{
	case CDDS_ITEMPREPAINT:
	{
		if (1 == nmCustomDraw.lItemlParam)
		{
			pNMCD->clrTextBk = RGB(0, 255, 0);
			pNMCD->clrText = RGB(0, 0, 0);
		}
		else if (2 == nmCustomDraw.lItemlParam)
		{
			pNMCD->clrTextBk = RGB(255, 0, 0);		//背景颜色
			pNMCD->clrText = RGB(0, 0, 0);		//文字颜色
		}
		else if (3 == nmCustomDraw.lItemlParam)
		{
			pNMCD->clrTextBk = RGB(200, 200, 200);
			pNMCD->clrText = RGB(128, 128, 128);
		}
		break;
	}
	default:
	{
		break;
	}
	}

	*pResult = 0;
	*pResult |= CDRF_NOTIFYPOSTPAINT;		//必须有，不然就没有效果
	*pResult |= CDRF_NOTIFYITEMDRAW;		//必须有，不然就没有效果
	return;
}


void CRJColorTestMSPGDlg::OnNMCustomdrawListTestresult(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMTVCUSTOMDRAW pNMCD = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);
	NMCUSTOMDRAW nmCustomDraw = pNMCD->nmcd;
	switch (nmCustomDraw.dwDrawStage)
	{
	case CDDS_ITEMPREPAINT:
	{
		if (1 == nmCustomDraw.lItemlParam)
		{
			pNMCD->clrTextBk = RGB(0, 255, 0);
			pNMCD->clrText = RGB(0, 0, 0);
		}
		else if (2 == nmCustomDraw.lItemlParam)
		{
			pNMCD->clrTextBk = RGB(255, 50, 50);		//背景颜色
			pNMCD->clrText = RGB(0, 0, 0);		//文字颜色
		}
		else if (3 == nmCustomDraw.lItemlParam)
		{
			pNMCD->clrTextBk = RGB(200, 200, 200);
			pNMCD->clrText = RGB(128, 128, 128);
		}
		break;
	}
	default:
	{
		break;
	}
	}

	*pResult = 0;
	*pResult |= CDRF_NOTIFYPOSTPAINT;		//必须有，不然就没有效果
	*pResult |= CDRF_NOTIFYITEMDRAW;		//必须有，不然就没有效果
	return;
}


void CRJColorTestMSPGDlg::CA_Measure_SxSyLv()
{
	float fLv, fx, fy;
	long lT;
	//float fduv;
	for (int i = 0; i < 4; i++) {
		m_pCaObj->Measure(0);
		fLv = m_pProbeObj->Lv;
		fx = m_pProbeObj->sx;
		fy = m_pProbeObj->sy;
		lT = m_pProbeObj->T;
		//fduv = m_pProbeObj->duv;

		m_fx.Format(_T("%1.3f"), fx);
		m_fy.Format(_T("%1.3f"), fy);
		m_Lv.Format(_T("%.2f"), fLv);
		m_CCT.Format(_T("%5ld"), lT);

		SetDlgItemText(IDC_EDIT_fx, m_fx);
		SetDlgItemText(IDC_EDIT_fy, m_fy);
		SetDlgItemText(IDC_EDIT_Lv, m_Lv);
		SetDlgItemText(IDC_EDIT_CCT, m_CCT);

		str_fx = CStringA(m_fx);
		str_fy = CStringA(m_fy);
		str_Lv = CStringA(m_Lv);
		str_CCT = CStringA(m_CCT);
		//UpdateData(FALSE);
	}
}


CString CRJColorTestMSPGDlg::timeChangeFormat()
{
	time_t rawtime;
	struct tm timeinfo;
	char s[100];
	time(&rawtime);
	localtime_s(&timeinfo, &rawtime);
	strftime(s, sizeof(s), "%Y%m%d-%H-%M-%S", &timeinfo);

	string strTime = s;
	CString strerr(strTime.c_str());
	return strerr;
}



void CRJColorTestMSPGDlg::OnBnClickedOk()
{
	if (m_pCaObj != NULL) {
		m_pCaObj->RemoteMode = 0;
	}

	outUserTestFile.close();


	CDialog::OnOK();
}

void CRJColorTestMSPGDlg::OnBnClickedCancel()
{
	if (m_pCaObj != NULL) {
		m_pCaObj->RemoteMode = 0;
	}

	outUserTestFile.close();

	CDialog::OnCancel();
}


BOOL CRJColorTestMSPGDlg::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN)
	{
		switch (pMsg->wParam)
		{
		case VK_RETURN://屏蔽回车
			return TRUE;
		case VK_ESCAPE://屏蔽Esc
			return TRUE;
		}
	}
	return CDialog::PreTranslateMessage(pMsg);
}


//获取可用的串口
void CRJColorTestMSPGDlg::OnBnClickedBtnFindcom()
{
	//先清空下拉框列表
	m_ComNumber.SetCurSel(-1);
	m_ComNumber.ResetContent();

	DWORD dwError = 0;
	CString strError = _T("");

	//打开注册表子键
	HKEY hRegKey;
	LONG lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _T("HARDWARE\\DEVICEMAP\\SERIALCOMM"), 0, KEY_READ, &hRegKey);  //只要读的权限就行
	if (lResult != ERROR_SUCCESS) {
		dwError = GetLastError();
		strError.Format(_T("打开COM注册表出错: 0x%08x \r\n"), dwError);
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(strError);
		return;
	}
	else {
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(_T("打开COM注册表成功！！！\r\n"));
	}

	//检索指定的子键下有多少个值项
	DWORD ValuesNumber;	
	lResult = RegQueryInfoKey(hRegKey, NULL, NULL, NULL, NULL, NULL, NULL, &ValuesNumber, NULL, NULL, NULL, NULL);
	if (lResult != ERROR_SUCCESS) {
		dwError = GetLastError();
		strError.Format(_T("检索连接在PC上的串口数量出错: 0x%08x \r\n"), dwError);
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(strError);
		RegCloseKey(hRegKey);
		return;
	}

	//检索每个值项，获取串口名、串口号
	DWORD index;
	WCHAR ValueName[100];
	DWORD ValueNameSize;
	BYTE DataBuffer[100];
	DWORD DataLen;
	CString strCOM;
	char temp;
	for (index = 0; index < ValuesNumber; index++)
	{
		memset(ValueName, 0, sizeof(ValueName));
		memset(DataBuffer, 0, sizeof(DataBuffer));
		ValueNameSize = 100;
		DataLen = 100;
		lResult = RegEnumValue(hRegKey, index, ValueName, &ValueNameSize, NULL, NULL, DataBuffer, &DataLen);
		if (lResult == ERROR_SUCCESS) {
			strCOM = _T("");
			for (DWORD i = 0; i < DataLen; i++) {
				temp = (char)DataBuffer[i];
				if (temp != '\0') {
					strCOM = strCOM + temp;
				}
			}
			m_ComNumber.AddString(strCOM);  //将识别到的串口 添加到下拉类表
		}
		else if (lResult == ERROR_NO_MORE_ITEMS) {
			//pMainDlg->AddToInfOut(_T("检索串口完毕！！！"),1,1);
		}
		else {
			dwError = GetLastError();
			strError.Format(_T("检索串口出错: 0x%08x \r\n"), dwError);
			m_EditShowC.SetSel(-1);
			m_EditShowC.ReplaceSel(strError);
			return;
		}
	}

	m_EditShowC.SetSel(-1);
	m_EditShowC.ReplaceSel(_T("获取串口列表成功！！！\r\n"));
	RegCloseKey(hRegKey);
	EnumComsName();//枚举所以可用串口的名称，显示在窗口，供参考
}

//枚举所以可用串口的名称，显示在窗口，供参考
void CRJColorTestMSPGDlg::EnumComsName()
{
	m_EditShowC.SetSel(-1);
	m_EditShowC.ReplaceSel(_T("********************************************\r\n"));
	m_EditShowC.ReplaceSel(_T("可用串口有:\r\n"));

	GUID* guidDev = (GUID*)&GUID_CLASS_COMPORT;
	HDEVINFO hDevInfo = SetupDiGetClassDevs(guidDev, NULL, NULL, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

	if (hDevInfo)
	{
		SP_DEVINFO_DATA SpDevInfo = { sizeof(SP_DEVINFO_DATA) };
		for (DWORD index = 0; SetupDiEnumDeviceInfo(hDevInfo, index, &SpDevInfo); index++)
		{
			//MessageBox((LPCTSTR)SpDevInfo.cbSize);
			TCHAR szName[512] = { 0 };
			if (SetupDiGetDeviceRegistryProperty(hDevInfo, &SpDevInfo, SPDRP_FRIENDLYNAME,
				NULL, (PBYTE)szName, sizeof(szName), NULL))
			{
				CString strComAllName, strComName, strComOpenName;
				strComAllName.Format(_T("%s"), szName);
				int posBeg = strComAllName.Find('(');
				int posEnd = strComAllName.Find(')');
				strComName = strComAllName.Mid(posBeg + 1, posEnd - posBeg - 1);
				strComOpenName = strComName + _T(":  ");
				//MessageBox(strComAllName);
				//m_Combo.AddString(strComOpenName + strComAllName);
				m_EditShowC.SetSel(-1);
				m_EditShowC.ReplaceSel(strComAllName);
				m_EditShowC.ReplaceSel(_T("\r\n"));
			}
		}
	}
	m_EditShowC.ReplaceSel(_T("********************************************\r\n"));
}



void CRJColorTestMSPGDlg::OnBnClickedBtnOpencom()
{
	CString cstr_COM, cstr_openCOM;
	m_ComNumber.GetWindowText(cstr_COM);
	cstr_openCOM = _T("\\\\.\\") + cstr_COM;  //需要用特殊格式打开串口，其中这个格式支持打开10以上的串口

	if (ComIsOK == false) {
		hCom = CreateFile(cstr_openCOM, //串口号
			GENERIC_READ | GENERIC_WRITE, //允许读或写
			0, //独占方式
			NULL,
			OPEN_EXISTING, //打开而不是创建
			FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED,//重叠方式,用于异步通信
			NULL);

		if (hCom == INVALID_HANDLE_VALUE)
		{
			AfxMessageBox(_T("打开COM失败，串口不存在或已被占用!"));
			ComIsOK = false;
			return;
		}
		//打开串口成功，ComIsOK 设为true; 串口选择下拉列表要置灰；该按钮文本改为关闭串口
		CString cstr_openCOMok = _T("打开串口") + cstr_COM + _T("成功!\r\n");
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(cstr_openCOMok);
		ComIsOK = true;
		HWND hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_Com);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_BaudRate);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_DataBit);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_StopBit);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_CheckBit);
		::EnableWindow(hBtn, FALSE);
		//打开连接和关闭7800的控件
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_Connect7800);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_Close7800);
		::EnableWindow(hBtn, TRUE);
		GetDlgItem(IDC_BTN_OpenCOM)->SetWindowTextW(_T("关闭串口"));
	}
	else {
		//第二次点击的时候，实现关闭串口的功能，串口选择下拉列表恢复；该按钮文本改为打开串口
		CloseHandle(hCom);//关闭句柄
		CString cstr_closeCOMok = _T("关闭串口") + cstr_COM + _T("成功~\r\n");
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(cstr_closeCOMok);
		ComIsOK = false;
		HWND hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_Com);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_BaudRate);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_DataBit);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_StopBit);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_COMBO_CheckBit);
		::EnableWindow(hBtn, TRUE);
		//置灰 连接和关闭7800的控件
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_Connect7800);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_Close7800);
		::EnableWindow(hBtn, FALSE);
		GetDlgItem(IDC_BTN_OpenCOM)->SetWindowTextW(_T("打开串口"));
		return;
	}

	//串口属性配置
	SetCommMask(hCom, EV_TXEMPTY | EV_RXCHAR); //设置事件掩码,暂时没用上
	SetupComm(hCom, 1024, 1024); //设置输入缓冲区和输出缓冲区的大小都是1024
	PurgeComm(hCom, PURGE_TXABORT | PURGE_RXABORT | PURGE_TXCLEAR | PURGE_RXCLEAR);//清理串口 	

	COMMTIMEOUTS TimeOuts; //设定读超时
	TimeOuts.ReadIntervalTimeout = MAXDWORD;
	TimeOuts.ReadTotalTimeoutConstant = 0;
	TimeOuts.ReadTotalTimeoutMultiplier = 0; //设定写超时
	TimeOuts.WriteTotalTimeoutConstant = 5000;
	TimeOuts.WriteTotalTimeoutMultiplier = 1000;
	if (SetCommTimeouts(hCom, &TimeOuts) == false)
	{
		CloseHandle(hCom);
		ComIsOK = false;
		AfxMessageBox(_T("串口超时设置失败!"));
		return;
	}

	//串口属性配置
	DCB dcb;
	GetCommState(hCom, &dcb);
	CString CStrComPropetry;
	GetDlgItem(IDC_COMBO_BaudRate)->GetWindowText(CStrComPropetry);
	dcb.BaudRate = _ttoi(CStrComPropetry); //dcb.BaudRate=9600; //波特率为9600
	GetDlgItem(IDC_COMBO_DataBit)->GetWindowText(CStrComPropetry);
	dcb.ByteSize = _ttoi(CStrComPropetry); //dcb.ByteSize=8; //每个字节为8位
	GetDlgItem(IDC_COMBO_StopBit)->GetWindowText(CStrComPropetry);
	dcb.StopBits = _ttoi(CStrComPropetry);  //dcb.StopBits=ONESTOPBIT;   //1位停止位
	GetDlgItem(IDC_COMBO_CheckBit)->GetWindowText(CStrComPropetry);
	dcb.Parity = _ttoi(CStrComPropetry);  //dcb.Parity=NOPARITY; //无奇偶校验位
	SetCommState(hCom, &dcb);
	PurgeComm(hCom, PURGE_TXCLEAR | PURGE_RXCLEAR);
	if (SetCommState(hCom, &dcb) == false)
	{
		CloseHandle(hCom);
		ComIsOK = false;
		AfxMessageBox(_T("串口属性设置失败!"));
		return;
	}

	return;
}


void CRJColorTestMSPGDlg::OnBnClickedBtnConnect7800()
{
	BYTE data[] = { 0x05, '\0' };
	int dataLen = sizeof(data) - 1;
	sendCommand(data, dataLen);

	HWND hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendTimingC);
	::EnableWindow(hBtn, TRUE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendPatternC);
	::EnableWindow(hBtn, TRUE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendUserC);
	::EnableWindow(hBtn, TRUE);

	if (CAIsOK == false) {
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ConnectCA);
		::EnableWindow(hBtn, TRUE);
	}
	else {
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_GammaTest);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ColorTest);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ContrastTest);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_AllTest);
		::EnableWindow(hBtn, TRUE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_UserTest);
		::EnableWindow(hBtn, TRUE);
	}
}


void CRJColorTestMSPGDlg::OnBnClickedBtnClose7800()
{
	BYTE data[] = { 0x04, '\0' };
	int dataLen = sizeof(data) - 1;
	sendCommand(data, dataLen);

	HWND hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendTimingC);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendPatternC);
	::EnableWindow(hBtn, FALSE);
	hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_SendUserC);
	::EnableWindow(hBtn, FALSE);

	if (CAIsOK == false) {
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ConnectCA);
		::EnableWindow(hBtn, FALSE);
	}
	else {
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_GammaTest);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ColorTest);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_ContrastTest);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_AllTest);
		::EnableWindow(hBtn, FALSE);
		hBtn = ::GetDlgItem(m_hWnd, IDC_BTN_UserTest);
		::EnableWindow(hBtn, FALSE);
	}
}

//监控下拉列表，监控7800播放时间的设置
void CRJColorTestMSPGDlg::OnCbnSelchangeCombo7800step()
{
	int curSel_steptime = vCombo_7800Step.GetCurSel();
	CString curStr_7800Step;
	vCombo_7800Step.GetLBText(curSel_steptime, curStr_7800Step);
	step7800 = _ttoi(curStr_7800Step);
}

void CRJColorTestMSPGDlg::OnBnClickedBtnSendtimingc()
{
	BYTE data[] = { 0x02,0x07,'\0' };
	int dataLen = sizeof(data) - 1;
	sendCommand(data, dataLen);

	CString str;
	GetDlgItem(IDC_EDIT_Timing)->GetWindowText(str);
	int HexLen;
	BYTE Hexbuf[1024];
	CStringtoBYTE(str, str.GetLength(), Hexbuf, &HexLen);
	sendCommand(Hexbuf, HexLen);

	BYTE data1[] = { 0x03,'\0' };
	int dataLen1 = sizeof(data1) - 1;
	sendCommand(data1, dataLen1);
}


void CRJColorTestMSPGDlg::OnBnClickedBtnSendpatternc()
{
	BYTE data[] = { 0x02,0x08,'\0' };
	int dataLen = sizeof(data) - 1;
	sendCommand(data, dataLen);

	CString str;
	GetDlgItem(IDC_EDIT_Pattern)->GetWindowText(str);
	int HexLen;
	BYTE Hexbuf[1024];
	CStringtoBYTE(str, str.GetLength(), Hexbuf, &HexLen);
	sendCommand(Hexbuf, HexLen);

	BYTE data1[] = { 0x03,'\0' };
	int dataLen1 = sizeof(data1) - 1;
	sendCommand(data1, dataLen1);
}


void CRJColorTestMSPGDlg::OnBnClickedBtnSenduserc()
{
	CString userCommand;
	GetDlgItem(IDC_EDIT_UserC)->GetWindowText(userCommand);//从编辑框控件中获取命令

	int HexLen;
	BYTE Hexbuf[1024];
	int strLen = userCommand.GetLength();

	CStringtoHexBYTE(userCommand, strLen, Hexbuf, &HexLen);
	sendCommand(Hexbuf, HexLen);
}


void CRJColorTestMSPGDlg::sendCommand(BYTE* pBHex, int HexLen)
{
	if (ComIsOK == false) {
		AfxMessageBox(_T("请先打开串口!"));
		return;
	}

	OVERLAPPED overlappedWrite;
	ZeroMemory(&overlappedWrite, sizeof(OVERLAPPED));
	overlappedWrite.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

	DWORD dwWritenSize = 0;
	PurgeComm(hCom, PURGE_TXABORT | PURGE_RXABORT | PURGE_TXCLEAR | PURGE_RXCLEAR);//清理串口 	
	BOOL bRet = ::WriteFile(hCom, pBHex, HexLen, &dwWritenSize, &overlappedWrite);


	if (!bRet)
	{
		if (ERROR_IO_PENDING == GetLastError()) {  //如果ReadFile挂起了，等待操作执行完
			while (!bRet)
			{
				bRet = GetOverlappedResult(hCom, &overlappedWrite, &dwWritenSize, TRUE); //最后参数为TRUE，操作完成之前不会有返回
				if (GetLastError() != ERROR_IO_INCOMPLETE) {
					break;
				}
			}
		}
	}

	CString IsSendOK;
	CString text = _T("->  ");
	CString textSend;

	if (bRet) {
		IsSendOK = _T(" - 发送成功!\r\n");
		textSend = BTYEtoCString(pBHex, HexLen);
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(text);
		m_EditShowC.ReplaceSel(textSend);
		m_EditShowC.ReplaceSel(IsSendOK);
	}
	else {
		
		IsSendOK = _T(" - 发送失败!\r\n");
		textSend = BTYEtoCString(pBHex, HexLen);
		m_EditShowC.SetSel(-1);
		m_EditShowC.ReplaceSel(text);
		m_EditShowC.ReplaceSel(textSend);
		m_EditShowC.ReplaceSel(IsSendOK);
	}

	PurgeComm(hCom, PURGE_TXABORT | PURGE_RXABORT | PURGE_TXCLEAR | PURGE_RXCLEAR);//清理串口 
	//关闭overlappedWrite事件
	if (overlappedWrite.hEvent != 0) {
		CloseHandle(overlappedWrite.hEvent);
	}

}


CString CRJColorTestMSPGDlg::BTYEtoCString(BYTE* BHex, int Hexlen)
{
	CString cstr = _T("");
	char temp;
	for (int i = 0; i < Hexlen; i++) {
		if (BHex[i] >= 0x00 && BHex[i] <= 0x09) {
			temp = BHex[i] + 0x30;
			cstr = cstr + _T("0") + temp;
		}
		else if (BHex[i] >= 0x0A && BHex[i] <= 0x0F) {
			temp = BHex[i] + 0x37;
			cstr = cstr + _T("0") + temp;
		}
		else if (BHex[i] >= 0x10 && BHex[i] <= 0x19) {
			temp = BHex[i] + 0x20;
			cstr = cstr + _T("1") + temp;
		}
		else if (BHex[i] >= 0x30 && BHex[i] <= 0x39) {
			temp = BHex[i];
			cstr = cstr + _T("3") + temp;
		}
		cstr = cstr + _T(" ");
	}
	return cstr;
}

void CRJColorTestMSPGDlg::CStringtoBYTE(CString str, int strLen, BYTE* BHex, int* Hexlen)
{
	*Hexlen = 0;
	for (int i = 0; i < strLen; i++) {
		if (str[i] < '0' || str[i] > '9') {
			continue;
		}
		BHex[*Hexlen] = (BYTE)str[i];
		(*Hexlen)++;
	}
}

void CRJColorTestMSPGDlg::CStringtoHexBYTE(CString str, int strLen, BYTE* BHex, int* Hexlen)
{
	*Hexlen = 0; //输出的16进制字符串长度
	char temp1, temp2; //接收一个字节的两个字符 例如EB，则temp1='E',temp2 = 'B'
	int j = 0; //用于两个字符分别赋值给temp1, temp2
	for (int i = 0; i < strLen; i++) {
		if (str[i] == ' ') {  //如果是空格和转义字符就跳过
			continue;
		}
		else {
			if (j == 0)
				temp1 = (char)str[i];
			if (j == 1)
				temp2 = (char)str[i];
			j++;
			if (j == 2) {
				j = 0;
				temp1 = temp1 - '0';
				temp2 = temp2 - '0';
				if (temp1 > 10)
					temp1 = temp1 - 7;
				if (temp2 > 10)
					temp2 = temp2 - 7;

				BHex[*Hexlen] = temp1 * 16 + temp2;
				(*Hexlen)++;
			}
		}
	}
}


