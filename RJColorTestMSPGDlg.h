
// RJColorTestMSPGDlg.h: 头文件
//

#pragma once

// CA - SDK
#import "C:\Program Files (x86)\KONICAMINOLTA\CA-SDK\SDK\CA200Srvr.dll" no_namespace no_implementation 


// CRJColorTestMSPGDlg 对话框
class CRJColorTestMSPGDlg : public CDialogEx
{
// 构造
public:
	CRJColorTestMSPGDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_RJCOLORTEST_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

	BOOL PreTranslateMessage(MSG* pMsg);//屏蔽回车和Esc

// 实现
protected:
	HICON m_hIcon;
	CBrush backBrush;

	//CA-SDK 智能指针类型
	IOutputProbesPtr m_pOutputProbesObj;
	IProbesPtr m_pProbesObj;
	ICasPtr m_pCasObj;
	IMemoryPtr m_pMemoryObj;
	IProbePtr m_pProbeObj;
	ICaPtr m_pCaObj;
	ICa200Ptr m_pCa200Obj;

	

public://串口&7800相关控件变量
	CComboBox m_ComNumber;
	CEdit m_EditShowC;
	CComboBox vCombo_7800Step;

public:
	HANDLE hCom = NULL; //串口句柄
	bool ComIsOK = false; //串口是否打开
	bool CAIsOK = false;  //CA是否连接上

public://CA310 相关控件变量
	CListCtrl m_isOpencaOK;
	CListCtrl m_isTestOver;
	CString m_fx;
	CString m_fy;
	CString m_Lv;
	CString m_CCT;

public://CA310 & 色彩测试相关控件操作
	afx_msg void OnBnClickedBtnGammatest();
	afx_msg void OnBnClickedBtnColortest();
	afx_msg void OnBnClickedBtnContrasttest();
	afx_msg void OnBnClickedBtnAlltest();
	afx_msg void OnBnClickedBtnConnectca();

	afx_msg void OnNMCustomdrawListOpenca(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnNMCustomdrawListTestresult(NMHDR* pNMHDR, LRESULT* pResult);

	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedBtnUsertest();
	afx_msg void OnBnClickedCheckIssavedata();

public://CA310 & 色彩测试相关方法
	void CA_Measure_SxSyLv();    //CA测试函数
	CString timeChangeFormat();   //系统时间更改成标准时间格式
	static DWORD __stdcall ThreadFunc_Gamma(LPVOID lpParam);
	static DWORD __stdcall ThreadFunc_Color(LPVOID lpParam);
	static DWORD __stdcall ThreadFunc_Contrast(LPVOID lpParam);
		

public://串口&7800相关控件操作
	afx_msg void OnBnClickedBtnFindcom();
	afx_msg void OnBnClickedBtnOpencom();
	afx_msg void OnBnClickedBtnConnect7800();
	afx_msg void OnBnClickedBtnClose7800();
	afx_msg void OnBnClickedBtnSendtimingc();
	afx_msg void OnBnClickedBtnSendpatternc();
	afx_msg void OnBnClickedBtnSenduserc();

public://串口&7800命令相关方法
	bool CRJColorTestMSPGDlg::EnumComs(struct UartInfo** UartCom, LPDWORD UartComNumber);
	void CRJColorTestMSPGDlg::sendCommand(BYTE* pBHex, int HexLen);
	CString CRJColorTestMSPGDlg::BTYEtoCString(BYTE* BHex, int Hexlen);
	void CRJColorTestMSPGDlg::CStringtoBYTE(CString str, int strLen, BYTE* BHex, int* Hexlen);
	void CRJColorTestMSPGDlg::CStringtoHexBYTE(CString str, int strLen, BYTE* BHex, int* Hexlen);
	
	afx_msg void OnCbnSelchangeCombo7800step();
};


//用于存放串口名和数量
struct UartInfo
{
	DWORD UartNum;
	WCHAR UartName[20];
};