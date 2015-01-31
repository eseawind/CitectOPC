// CitectOPC.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <afxcoll.h>
#include <iostream>
#include "opccomn.h"
#include "opcda.h"

#include "opcerror.h"
#include "OPCEnum.h"

#include "opccomn_i.c"
#include "opcda_i.c"
#include "OPCEnum_i.c"
using namespace std;
//#include "COPCDataCallback.h"

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr;
	COSERVERINFO si;
	ZeroMemory(&si, sizeof(si)); //内存1 存放服务器信息
	MULTI_QI mqi[1];
	ZeroMemory(&mqi, sizeof(mqi)); //内存2 存放MULTI_QI用以获取OPC服务器列表
	CATID catID[2];
	ZeroMemory(&catID, 2 * sizeof(CATID));  //内存3 存放OPC标准接口1.0和2.0
	CLSID clsid[20];
	ZeroMemory(&clsid, 20 * sizeof(CLSID)); //内存4 存放枚举出的OPC服务器的CLSID

	/*Initialize DCOM*/
	CString strError;
	CString strNode = _T("192.168.1.11"); //The IP Address of the OPC SERVER
	CString strProgID = _T("Citect.OPC");  //The ProgID of the OPC SERVER	

	/*初始化指针*/
	IOPCServerList *pIServerList = NULL;  //第一个接口指针，指向OPC服务器列表
	IEnumGUID *pIEnumGUID = NULL; //第二个接口指针，指向OPC服务器列表枚举
	IUnknown *pIUnknown = NULL; //第三个指针，服务器接口指针
	IOPCServer *pIServer = NULL; //第四个指针，指向OPC服务器的接口指针
	IOPCItemMgt *pIOPCItemMgt = NULL; //第五个指针，指向标签的指针
	OPCITEMDEF *itemArray = NULL; //第六个指针，指向标签内容数组
	OPCITEMRESULT *pItemResult = NULL;  //第七个指针，指向添加项的结果
	HRESULT *pErrors = NULL; //第八个指针，指向添加项错误
	OPCHANDLE *hOPCServer2 = NULL; //第九个指针，指向第二个OPCServer句柄
	IOPCAsyncIO2 *pIOPCAsyncIO2 = NULL; /*第十一个指针，指向异步读取*/
	IOPCGroupStateMgt *pIOPCGroupStateMgt = NULL; //第十个指针，指向组状态更新

	//IConnectionPointContainer *pIConnectPointContainer = NULL;
	//IConnectionPoint *pIConnectionPoint = NULL;

	/*COM组件注册及登陆*/
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED); //注册DCOM组件，获取结果
	HRESULT hr_sec = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_NONE, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
	ASSERT(hr_sec == S_OK || RPC_E_TOO_LATE == hr_sec);

	if (FAILED(hr)) {
		cout << "未开启COM组件注册..." << endl;
		return 1;
	}
	else if (SUCCEEDED(hr)) {		
		cout << "已开启COM组件注册..." << endl;
		cout << "已登陆COM组件..." << endl;
	}

	/*获取OPC服务器列表*/	

	si.pwszName = (LPWSTR)(strNode.GetString()); //Covert CString to LPWSTR 把服务器名转为宽字节字符串	

	mqi[0].hr = S_OK;
	mqi[0].pIID = &IID_IOPCServerList;
	mqi[0].pItf = NULL;

	hr = CoCreateInstanceEx(CLSID_OpcServerList, NULL, CLSCTX_ALL, &si, 1, mqi); //Connect to the target OPC SERVER & Get the target SERVER's OPC SERVER List

	if (FAILED(hr)) {
		cout << "获取目标服务器端OPC服务器列表失败..." << endl;
	}
	else if (SUCCEEDED(hr)) {
		cout << "获取目标服务器端OPC服务器列表成功..." << endl;
	}

	/*枚举OPC服务器*/
	pIServerList = (IOPCServerList*)mqi[0].pItf; //得到第一个指针
	ASSERT(pIServerList);
	
	catID[0] = CATID_OPCDAServer10;
	catID[1] = CATID_OPCDAServer20;
	
	hr = pIServerList->EnumClassesOfCategories(2, catID, 2, catID, &pIEnumGUID); //Enum OPCDASERVER1.0 & OPCDASERVER2.0 得到第二个指针
	ASSERT(pIEnumGUID);

	if (FAILED(hr)) {
		cout << "未获取到OPCDAServer1.0和OPCDAServer2.0接口" << endl;
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {
		cout << "已获取到OPCDAServer1.0和OPCDAServer2.0接口" << endl;
	}

	ULONG nCount;
	
	HRESULT hr_1;
	cout << "搜索到的接口列表：" << endl;
	do
	{
		hr_1 = pIEnumGUID->Next(20, clsid, &nCount);
		for (ULONG i = 0; i < nCount; i++)
		{
			LPOLESTR szProID;
			LPOLESTR szUserType;
			HRESULT hr_2 = pIServerList->GetClassDetails(clsid[i], &szProID, &szUserType);
			ASSERT(hr_2 == S_OK);
			wcout << szProID << endl;
			CoTaskMemFree(szProID);
			CoTaskMemFree(szUserType);
		}
	} while (hr_1 == S_OK);
	cout << "枚举完毕..." << endl;

	/*连接本地或远程OPC服务器*/
	CLSID clsid_citect;

	hr = pIServerList->CLSIDFromProgID(strProgID, &clsid_citect); //获取OPC服务器的CLSID

	if (FAILED(hr)) {
		cout << "未找到需要连接的OPC服务器..." << endl;
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {
		cout << "已找到需要连接的OPC服务器..." << endl;
	}

	mqi[0].hr = S_OK;
	mqi[0].pIID = &IID_IOPCServer;
	mqi[0].pItf = NULL;

	hr = CoCreateInstanceEx(clsid_citect, NULL, CLSCTX_ALL, &si, 1, mqi);

	if (FAILED(hr)) {
		cout << "未连接上需要连接的OPC服务器..." << endl;
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {
		cout << "已连接上需要连接的OPC服务器..." << endl;
	}

	/*获取OPC服务器IUnknow接口指针*/

	pIUnknown = (IUnknown*)mqi[0].pItf;  //得到第三个指针
	ASSERT(pIUnknown);
	
	LONG lTimeBias = 0;
	FLOAT fTemp = 0.00;
	OPCHANDLE hOPCServer1; //第一个OPC SERVER句柄
	DWORD dwActualRate = 0;

	hr = pIUnknown->QueryInterface(IID_IOPCServer,/*OUT*/(void**)&pIServer); //得到第四个指针
	ASSERT(pIServer);

	if (FAILED(hr)){
		cout << "未能获取IOPCServer接口..." << endl;
		CoTaskMemFree(&hOPCServer1);
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIServer) pIServer->Release(); //第四个内存释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个内存释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr))
	{		
		cout << "已获取IOPCServer接口" << endl;
	}

	/*为OPC服务器添加组*/

	/*潘河 为避免代码冗余，所有潘河数据全部放在一组，用Item标识区别*/

	hr = pIServer->AddGroup(L"PH", TRUE, 10, 1235, &lTimeBias, &fTemp, 0, /*OUT*/ &hOPCServer1, 
		&dwActualRate, IID_IOPCItemMgt, /*OUT*/ (LPUNKNOWN*)&pIOPCItemMgt); //得到第五个指针
	ASSERT(pIOPCItemMgt);
	ASSERT(hOPCServer1);

	if (FAILED(hr)) {
		cout << "未能添加潘河组..." << endl;
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
		pIOPCItemMgt = NULL;
		if (pIServer) pIServer->Release(); //第四个指针释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个指针释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {
		cout << "已添加潘河组..." << endl;
	}

	/*为潘河组添加条目*/

	itemArray = (OPCITEMDEF*)CoTaskMemAlloc(119 * sizeof(OPCITEMDEF)); //第五个内存 指向存放数据项

	/* PH63_10阀组（2#）*/

	/*汇管温度*/
	CString strPH63HT = "PH1_VALUE_PH63_HT";
	BSTR bstrPH63HT = strPH63HT.AllocSysString();
	itemArray[0].szAccessPath = L"";
	itemArray[0].szItemID = bstrPH63HT;
	itemArray[0].bActive = true; // active state
	itemArray[0].hClient = (OPCHANDLE)0; // our handle to item
	itemArray[0].dwBlobSize = 0; // no blob support
	itemArray[0].pBlob = NULL;
	itemArray[0].vtRequestedDataType = VT_R4;
	itemArray[0].wReserved = 0;

	/*汇管压力*/
	CString strPH63HP = "PH1_VALUE_PH63_HP";
	BSTR bstrPH63HP = strPH63HP.AllocSysString();
	itemArray[1].szAccessPath = L"";
	itemArray[1].szItemID = bstrPH63HP;
	itemArray[1].bActive = true; // active state
	itemArray[1].hClient = (OPCHANDLE)1; // our handle to item
	itemArray[1].dwBlobSize = 0; // no blob support
	itemArray[1].pBlob = NULL;
	itemArray[1].vtRequestedDataType = VT_R4;
	itemArray[1].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH63HF = "PH1_VALUE_PH63_HF";
	BSTR bstrPH63HF = strPH63HF.AllocSysString();
	itemArray[2].szAccessPath = L"";
	itemArray[2].szItemID = bstrPH63HF;
	itemArray[2].bActive = true; // active state
	itemArray[2].hClient = (OPCHANDLE)2; // our handle to item
	itemArray[2].dwBlobSize = 0; // no blob support
	itemArray[2].pBlob = NULL;
	itemArray[2].vtRequestedDataType = VT_R4;
	itemArray[2].wReserved = 0;

	/*汇管累计流量*/
	CString strPH63HS = "PH1_VALUE_PH63_HS";
	BSTR bstrPH63HS = strPH63HS.AllocSysString();
	itemArray[3].szAccessPath = L"";
	itemArray[3].szItemID = bstrPH63HS;
	itemArray[3].bActive = true; // active state
	itemArray[3].hClient = (OPCHANDLE)3; // our handle to item
	itemArray[3].dwBlobSize = 0; // no blob support
	itemArray[3].pBlob = NULL;
	itemArray[3].vtRequestedDataType = VT_R4;
	itemArray[3].wReserved = 0;

	/*温度*/
	CString strPH63DT = "PH1_VALUE_PH63_DT";
	BSTR bstrPH63DT = strPH63DT.AllocSysString();
	itemArray[4].szAccessPath = L"";
	itemArray[4].szItemID = bstrPH63DT;
	itemArray[4].bActive = true; // active state
	itemArray[4].hClient = (OPCHANDLE)4; // our handle to item
	itemArray[4].dwBlobSize = 0; // no blob support
	itemArray[4].pBlob = NULL;
	itemArray[4].vtRequestedDataType = VT_R4;
	itemArray[4].wReserved = 0;

	/*压力*/
	CString strPH63DP = "PH1_VALUE_PH63_DP";
	BSTR bstrPH63DP = strPH63DP.AllocSysString();
	itemArray[5].szAccessPath = L"";
	itemArray[5].szItemID = bstrPH63DP;
	itemArray[5].bActive = true; // active state
	itemArray[5].hClient = (OPCHANDLE)5; // our handle to item
	itemArray[5].dwBlobSize = 0; // no blob support
	itemArray[5].pBlob = NULL;
	itemArray[5].vtRequestedDataType = VT_R4;
	itemArray[5].wReserved = 0;

	/*瞬时流量*/
	CString strPH63DF = "PH1_VALUE_PH63_DF";
	BSTR bstrPH63DF = strPH63DF.AllocSysString();
	itemArray[6].szAccessPath = L"";
	itemArray[6].szItemID = bstrPH63DF;
	itemArray[6].bActive = true; // active state
	itemArray[6].hClient = (OPCHANDLE)6; // our handle to item
	itemArray[6].dwBlobSize = 0; // no blob support
	itemArray[6].pBlob = NULL;
	itemArray[6].vtRequestedDataType = VT_R4;
	itemArray[6].wReserved = 0;

	/*累计流量*/
	CString strPH63DS = "PH1_VALUE_PH63_DS";
	BSTR bstrPH63DS = strPH63DS.AllocSysString();
	itemArray[7].szAccessPath = L"";
	itemArray[7].szItemID = bstrPH63DS;
	itemArray[7].bActive = true; // active state
	itemArray[7].hClient = (OPCHANDLE)7; // our handle to item
	itemArray[7].dwBlobSize = 0; // no blob support
	itemArray[7].pBlob = NULL;
	itemArray[7].vtRequestedDataType = VT_R4;
	itemArray[7].wReserved = 0;

	/* PH73_11阀组（3#）*/

	/*汇管温度*/
	CString strPH7311HT = "PH1_VALUE_PH7311_HT";
	BSTR bstrPH7311HT = strPH7311HT.AllocSysString();
	itemArray[8].szAccessPath = L"";
	itemArray[8].szItemID = bstrPH7311HT;
	itemArray[8].bActive = true; // active state
	itemArray[8].hClient = (OPCHANDLE)8; // our handle to item
	itemArray[8].dwBlobSize = 0; // no blob support
	itemArray[8].pBlob = NULL;
	itemArray[8].vtRequestedDataType = VT_R4;
	itemArray[8].wReserved = 0;

	/*汇管压力*/
	CString strPH7311HP = "PH1_VALUE_PH7311_HP";
	BSTR bstrPH7311HP = strPH7311HP.AllocSysString();
	itemArray[9].szAccessPath = L"";
	itemArray[9].szItemID = bstrPH7311HP;
	itemArray[9].bActive = true; // active state
	itemArray[9].hClient = (OPCHANDLE)9; // our handle to item
	itemArray[9].dwBlobSize = 0; // no blob support
	itemArray[9].pBlob = NULL;
	itemArray[9].vtRequestedDataType = VT_R4;
	itemArray[9].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH7311HF = "PH1_VALUE_PH7311_HF";
	BSTR bstrPH7311HF = strPH7311HF.AllocSysString();
	itemArray[10].szAccessPath = L"";
	itemArray[10].szItemID = bstrPH7311HF;
	itemArray[10].bActive = true; // active state
	itemArray[10].hClient = (OPCHANDLE)10; // our handle to item
	itemArray[10].dwBlobSize = 0; // no blob support
	itemArray[10].pBlob = NULL;
	itemArray[10].vtRequestedDataType = VT_R4;
	itemArray[10].wReserved = 0;

	/*汇管累计流量*/
	CString strPH7311HS = "PH1_VALUE_PH7311_HS";
	BSTR bstrPH7311HS = strPH7311HS.AllocSysString();
	itemArray[11].szAccessPath = L"";
	itemArray[11].szItemID = bstrPH7311HS;
	itemArray[11].bActive = true; // active state
	itemArray[11].hClient = (OPCHANDLE)11; // our handle to item
	itemArray[11].dwBlobSize = 0; // no blob support
	itemArray[11].pBlob = NULL;
	itemArray[11].vtRequestedDataType = VT_R4;
	itemArray[11].wReserved = 0;

	/*温度*/
	CString strPH7311DT = "PH1_VALUE_PH7311_DT";
	BSTR bstrPH7311DT = strPH7311DT.AllocSysString();
	itemArray[12].szAccessPath = L"";
	itemArray[12].szItemID = bstrPH7311DT;
	itemArray[12].bActive = true; // active state
	itemArray[12].hClient = (OPCHANDLE)12; // our handle to item
	itemArray[12].dwBlobSize = 0; // no blob support
	itemArray[12].pBlob = NULL;
	itemArray[12].vtRequestedDataType = VT_R4;
	itemArray[12].wReserved = 0;

	/*压力*/
	CString strPH7311DP = "PH1_VALUE_PH7311_DP";
	BSTR bstrPH7311DP = strPH7311DP.AllocSysString();
	itemArray[13].szAccessPath = L"";
	itemArray[13].szItemID = bstrPH7311DP;
	itemArray[13].bActive = true; // active state
	itemArray[13].hClient = (OPCHANDLE)13; // our handle to item
	itemArray[13].dwBlobSize = 0; // no blob support
	itemArray[13].pBlob = NULL;
	itemArray[13].vtRequestedDataType = VT_R4;
	itemArray[13].wReserved = 0;

	/*瞬时流量*/
	CString strPH7311DF = "PH1_VALUE_PH7311_DF";
	BSTR bstrPH7311DF = strPH7311DF.AllocSysString();
	itemArray[14].szAccessPath = L"";
	itemArray[14].szItemID = bstrPH7311DF;
	itemArray[14].bActive = true; // active state
	itemArray[14].hClient = (OPCHANDLE)14; // our handle to item
	itemArray[14].dwBlobSize = 0; // no blob support
	itemArray[14].pBlob = NULL;
	itemArray[14].vtRequestedDataType = VT_R4;
	itemArray[14].wReserved = 0;

	/*累计流量*/
	CString strPH7311DS = "PH1_VALUE_PH7311_DS";
	BSTR bstrPH7311DS = strPH7311DS.AllocSysString();
	itemArray[15].szAccessPath = L"";
	itemArray[15].szItemID = bstrPH7311DS;
	itemArray[15].bActive = true; // active state
	itemArray[15].hClient = (OPCHANDLE)15; // our handle to item
	itemArray[15].dwBlobSize = 0; // no blob support
	itemArray[15].pBlob = NULL;
	itemArray[15].vtRequestedDataType = VT_R4;
	itemArray[15].wReserved = 0;

	/* PH74_08阀组（4#）*/

	/*汇管温度*/
	CString strPH74HT = "PH1_VALUE_PH74_HT";
	BSTR bstrPH74HT = strPH74HT.AllocSysString();
	itemArray[16].szAccessPath = L"";
	itemArray[16].szItemID = bstrPH74HT;
	itemArray[16].bActive = true; // active state
	itemArray[16].hClient = (OPCHANDLE)16; // our handle to item
	itemArray[16].dwBlobSize = 0; // no blob support
	itemArray[16].pBlob = NULL;
	itemArray[16].vtRequestedDataType = VT_R4;
	itemArray[16].wReserved = 0;

	/*汇管压力*/
	CString strPH74HP = "PH1_VALUE_PH74_HP";
	BSTR bstrPH74HP = strPH74HP.AllocSysString();
	itemArray[17].szAccessPath = L"";
	itemArray[17].szItemID = bstrPH74HP;
	itemArray[17].bActive = true; // active state
	itemArray[17].hClient = (OPCHANDLE)17; // our handle to item
	itemArray[17].dwBlobSize = 0; // no blob support
	itemArray[17].pBlob = NULL;
	itemArray[17].vtRequestedDataType = VT_R4;
	itemArray[17].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH74HF = "PH1_VALUE_PH74_HF";
	BSTR bstrPH74HF = strPH74HF.AllocSysString();
	itemArray[18].szAccessPath = L"";
	itemArray[18].szItemID = bstrPH74HF;
	itemArray[18].bActive = true; // active state
	itemArray[18].hClient = (OPCHANDLE)18; // our handle to item
	itemArray[18].dwBlobSize = 0; // no blob support
	itemArray[18].pBlob = NULL;
	itemArray[18].vtRequestedDataType = VT_R4;
	itemArray[18].wReserved = 0;

	/*汇管累计流量*/
	CString strPH74HS = "PH1_VALUE_PH74_HS";
	BSTR bstrPH74HS = strPH74HS.AllocSysString();
	itemArray[19].szAccessPath = L"";
	itemArray[19].szItemID = bstrPH74HS;
	itemArray[19].bActive = true; // active state
	itemArray[19].hClient = (OPCHANDLE)19; // our handle to item
	itemArray[19].dwBlobSize = 0; // no blob support
	itemArray[19].pBlob = NULL;
	itemArray[19].vtRequestedDataType = VT_R4;
	itemArray[19].wReserved = 0;

	/*温度*/
	CString strPH74DT = "PH1_VALUE_PH74_DT";
	BSTR bstrPH74DT = strPH74DT.AllocSysString();
	itemArray[20].szAccessPath = L"";
	itemArray[20].szItemID = bstrPH74DT;
	itemArray[20].bActive = true; // active state
	itemArray[20].hClient = (OPCHANDLE)20; // our handle to item
	itemArray[20].dwBlobSize = 0; // no blob support
	itemArray[20].pBlob = NULL;
	itemArray[20].vtRequestedDataType = VT_R4;
	itemArray[20].wReserved = 0;

	/*压力*/
	CString strPH74DP = "PH1_VALUE_PH74_DP";
	BSTR bstrPH74DP = strPH74DP.AllocSysString();
	itemArray[21].szAccessPath = L"";
	itemArray[21].szItemID = bstrPH74DP;
	itemArray[21].bActive = true; // active state
	itemArray[21].hClient = (OPCHANDLE)21; // our handle to item
	itemArray[21].dwBlobSize = 0; // no blob support
	itemArray[21].pBlob = NULL;
	itemArray[21].vtRequestedDataType = VT_R4;
	itemArray[21].wReserved = 0;

	/*瞬时流量*/
	CString strPH74DF = "PH1_VALUE_PH74_DF";
	BSTR bstrPH74DF = strPH74DF.AllocSysString();
	itemArray[22].szAccessPath = L"";
	itemArray[22].szItemID = bstrPH74DF;
	itemArray[22].bActive = true; // active state
	itemArray[22].hClient = (OPCHANDLE)22; // our handle to item
	itemArray[22].dwBlobSize = 0; // no blob support
	itemArray[22].pBlob = NULL;
	itemArray[22].vtRequestedDataType = VT_R4;
	itemArray[22].wReserved = 0;

	/*累计流量*/
	CString strPH74DS = "PH1_VALUE_PH74_DS";
	BSTR bstrPH74DS = strPH74DS.AllocSysString();
	itemArray[23].szAccessPath = L"";
	itemArray[23].szItemID = bstrPH74DS;
	itemArray[23].bActive = true; // active state
	itemArray[23].hClient = (OPCHANDLE)23; // our handle to item
	itemArray[23].dwBlobSize = 23; // no blob support
	itemArray[23].pBlob = NULL;
	itemArray[23].vtRequestedDataType = VT_R4;
	itemArray[23].wReserved = 0;

	/* PH64_01阀组（5#）*/

	/*汇管温度*/
	CString strPH64HT = "PH1_VALUE_PH64_HT";
	BSTR bstrPH64HT = strPH64HT.AllocSysString();
	itemArray[24].szAccessPath = L"";
	itemArray[24].szItemID = bstrPH64HT;
	itemArray[24].bActive = true; // active state
	itemArray[24].hClient = (OPCHANDLE)24; // our handle to item
	itemArray[24].dwBlobSize = 0; // no blob support
	itemArray[24].pBlob = NULL;
	itemArray[24].vtRequestedDataType = VT_R4;
	itemArray[24].wReserved = 0;

	/*汇管压力*/
	CString strPH64HP = "PH1_VALUE_PH64_HP";
	BSTR bstrPH64HP = strPH64HP.AllocSysString();
	itemArray[25].szAccessPath = L"";
	itemArray[25].szItemID = bstrPH64HP;
	itemArray[25].bActive = true; // active state
	itemArray[25].hClient = (OPCHANDLE)25; // our handle to item
	itemArray[25].dwBlobSize = 0; // no blob support
	itemArray[25].pBlob = NULL;
	itemArray[25].vtRequestedDataType = VT_R4;
	itemArray[25].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH64HF = "PH1_VALUE_PH64_HF";
	BSTR bstrPH64HF = strPH64HF.AllocSysString();
	itemArray[26].szAccessPath = L"";
	itemArray[26].szItemID = bstrPH64HF;
	itemArray[26].bActive = true; // active state
	itemArray[26].hClient = (OPCHANDLE)26; // our handle to item
	itemArray[26].dwBlobSize = 0; // no blob support
	itemArray[26].pBlob = NULL;
	itemArray[26].vtRequestedDataType = VT_R4;
	itemArray[26].wReserved = 0;

	/*汇管累计流量*/
	CString strPH64HS = "PH1_VALUE_PH64_HS";
	BSTR bstrPH64HS = strPH64HS.AllocSysString();
	itemArray[27].szAccessPath = L"";
	itemArray[27].szItemID = bstrPH64HS;
	itemArray[27].bActive = true; // active state
	itemArray[27].hClient = (OPCHANDLE)27; // our handle to item
	itemArray[27].dwBlobSize = 0; // no blob support
	itemArray[27].pBlob = NULL;
	itemArray[27].vtRequestedDataType = VT_R4;
	itemArray[27].wReserved = 0;

	/*温度*/
	CString strPH64DT = "PH1_VALUE_PH64_DT";
	BSTR bstrPH64DT = strPH64DT.AllocSysString();
	itemArray[28].szAccessPath = L"";
	itemArray[28].szItemID = bstrPH64DT;
	itemArray[28].bActive = true; // active state
	itemArray[28].hClient = (OPCHANDLE)28; // our handle to item
	itemArray[28].dwBlobSize = 0; // no blob support
	itemArray[28].pBlob = NULL;
	itemArray[28].vtRequestedDataType = VT_R4;
	itemArray[28].wReserved = 0;

	/*压力*/
	CString strPH64DP = "PH1_VALUE_PH64_DP";
	BSTR bstrPH64DP = strPH64DP.AllocSysString();
	itemArray[29].szAccessPath = L"";
	itemArray[29].szItemID = bstrPH64DP;
	itemArray[29].bActive = true; // active state
	itemArray[29].hClient = (OPCHANDLE)29; // our handle to item
	itemArray[29].dwBlobSize = 0; // no blob support
	itemArray[29].pBlob = NULL;
	itemArray[29].vtRequestedDataType = VT_R4;
	itemArray[29].wReserved = 0;

	/*瞬时流量*/
	CString strPH64DF = "PH1_VALUE_PH64_DF";
	BSTR bstrPH64DF = strPH64DF.AllocSysString();
	itemArray[30].szAccessPath = L"";
	itemArray[30].szItemID = bstrPH64DF;
	itemArray[30].bActive = true; // active state
	itemArray[30].hClient = (OPCHANDLE)30; // our handle to item
	itemArray[30].dwBlobSize = 0; // no blob support
	itemArray[30].pBlob = NULL;
	itemArray[30].vtRequestedDataType = VT_R4;
	itemArray[30].wReserved = 0;

	/*累计流量*/
	CString strPH64DS = "PH1_VALUE_PH64_DS";
	BSTR bstrPH64DS = strPH64DS.AllocSysString();
	itemArray[31].szAccessPath = L"";
	itemArray[31].szItemID = bstrPH64DS;
	itemArray[31].bActive = true; // active state
	itemArray[31].hClient = (OPCHANDLE)31; // our handle to item
	itemArray[31].dwBlobSize = 0; // no blob support
	itemArray[31].pBlob = NULL;
	itemArray[31].vtRequestedDataType = VT_R4;
	itemArray[31].wReserved = 0;

	/* P86_01阀组（7#）*/

	/*汇管温度*/
	CString strPH86HT = "PH1_VALUE_PH86_HT";
	BSTR bstrPH86HT = strPH86HT.AllocSysString();
	itemArray[32].szAccessPath = L"";
	itemArray[32].szItemID = bstrPH86HT;
	itemArray[32].bActive = true; // active state
	itemArray[32].hClient = (OPCHANDLE)32; // our handle to item
	itemArray[32].dwBlobSize = 0; // no blob support
	itemArray[32].pBlob = NULL;
	itemArray[32].vtRequestedDataType = VT_R4;
	itemArray[32].wReserved = 0;

	/*汇管压力*/
	CString strPH86HP = "PH1_VALUE_PH86_HP";
	BSTR bstrPH86HP = strPH86HP.AllocSysString();
	itemArray[33].szAccessPath = L"";
	itemArray[33].szItemID = bstrPH86HP;
	itemArray[33].bActive = true; // active state
	itemArray[33].hClient = (OPCHANDLE)33; // our handle to item
	itemArray[33].dwBlobSize = 0; // no blob support
	itemArray[33].pBlob = NULL;
	itemArray[33].vtRequestedDataType = VT_R4;
	itemArray[33].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH86HF = "PH1_VALUE_PH86_HF";
	BSTR bstrPH86HF = strPH86HF.AllocSysString();
	itemArray[34].szAccessPath = L"";
	itemArray[34].szItemID = bstrPH86HF;
	itemArray[34].bActive = true; // active state
	itemArray[34].hClient = (OPCHANDLE)34; // our handle to item
	itemArray[34].dwBlobSize = 0; // no blob support
	itemArray[34].pBlob = NULL;
	itemArray[34].vtRequestedDataType = VT_R4;
	itemArray[34].wReserved = 0;

	/*汇管累计流量*/
	CString strPH86HS = "PH1_VALUE_PH86_HS";
	BSTR bstrPH86HS = strPH86HS.AllocSysString();
	itemArray[35].szAccessPath = L"";
	itemArray[35].szItemID = bstrPH86HS;
	itemArray[35].bActive = true; // active state
	itemArray[35].hClient = (OPCHANDLE)35; // our handle to item
	itemArray[35].dwBlobSize = 0; // no blob support
	itemArray[35].pBlob = NULL;
	itemArray[35].vtRequestedDataType = VT_R4;
	itemArray[35].wReserved = 0;

	/*温度*/
	CString strPH86DT = "PH1_VALUE_PH86_DT";
	BSTR bstrPH86DT = strPH86DT.AllocSysString();
	itemArray[36].szAccessPath = L"";
	itemArray[36].szItemID = bstrPH86DT;
	itemArray[36].bActive = true; // active state
	itemArray[36].hClient = (OPCHANDLE)36; // our handle to item
	itemArray[36].dwBlobSize = 0; // no blob support
	itemArray[36].pBlob = NULL;
	itemArray[36].vtRequestedDataType = VT_R4;
	itemArray[36].wReserved = 0;

	/*压力*/
	CString strPH86DP = "PH1_VALUE_PH86_DP";
	BSTR bstrPH86DP = strPH86DP.AllocSysString();
	itemArray[37].szAccessPath = L"";
	itemArray[37].szItemID = bstrPH86DP;
	itemArray[37].bActive = true; // active state
	itemArray[37].hClient = (OPCHANDLE)37; // our handle to item
	itemArray[37].dwBlobSize = 0; // no blob support
	itemArray[37].pBlob = NULL;
	itemArray[37].vtRequestedDataType = VT_R4;
	itemArray[37].wReserved = 0;

	/*瞬时流量*/
	CString strPH86DF = "PH1_VALUE_PH86_DF";
	BSTR bstrPH86DF = strPH86DF.AllocSysString();
	itemArray[38].szAccessPath = L"";
	itemArray[38].szItemID = bstrPH86DF;
	itemArray[38].bActive = true; // active state
	itemArray[38].hClient = (OPCHANDLE)38; // our handle to item
	itemArray[38].dwBlobSize = 0; // no blob support
	itemArray[38].pBlob = NULL;
	itemArray[38].vtRequestedDataType = VT_R4;
	itemArray[38].wReserved = 0;

	/*累计流量*/
	CString strPH86DS = "PH1_VALUE_PH86_DS";
	BSTR bstrPH86DS = strPH86DS.AllocSysString();
	itemArray[39].szAccessPath = L"";
	itemArray[39].szItemID = bstrPH86DS;
	itemArray[39].bActive = true; // active state
	itemArray[39].hClient = (OPCHANDLE)39; // our handle to item
	itemArray[39].dwBlobSize = 0; // no blob support
	itemArray[39].pBlob = NULL;
	itemArray[39].vtRequestedDataType = VT_R4;
	itemArray[39].wReserved = 0;

	/* PH65_11阀组（8#）*/
	
	/*汇管温度*/
	CString strPH65HT = "PH1_VALUE_PH65_HT";
	BSTR bstrPH65HT = strPH65HT.AllocSysString();
	itemArray[40].szAccessPath = L"";
	itemArray[40].szItemID = bstrPH65HT;
	itemArray[40].bActive = true; // active state
	itemArray[40].hClient = (OPCHANDLE)40; // our handle to item
	itemArray[40].dwBlobSize = 0; // no blob support
	itemArray[40].pBlob = NULL;
	itemArray[40].vtRequestedDataType = VT_R4;
	itemArray[40].wReserved = 0;

	/*汇管压力*/
	CString strPH65HP = "PH1_VALUE_PH65_HP";
	BSTR bstrPH65HP = strPH65HP.AllocSysString();
	itemArray[41].szAccessPath = L"";
	itemArray[41].szItemID = bstrPH65HP;
	itemArray[41].bActive = true; // active state
	itemArray[41].hClient = (OPCHANDLE)41; // our handle to item
	itemArray[41].dwBlobSize = 0; // no blob support
	itemArray[41].pBlob = NULL;
	itemArray[41].vtRequestedDataType = VT_R4;
	itemArray[41].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH65HF = "PH1_VALUE_PH65_HF";
	BSTR bstrPH65HF = strPH65HF.AllocSysString();
	itemArray[42].szAccessPath = L"";
	itemArray[42].szItemID = bstrPH65HF;
	itemArray[42].bActive = true; // active state
	itemArray[42].hClient = (OPCHANDLE)42; // our handle to item
	itemArray[42].dwBlobSize = 0; // no blob support
	itemArray[42].pBlob = NULL;
	itemArray[42].vtRequestedDataType = VT_R4;
	itemArray[42].wReserved = 0;

	/*汇管累计流量*/
	CString strPH65HS = "PH1_VALUE_PH65_HS";
	BSTR bstrPH65HS = strPH65HS.AllocSysString();
	itemArray[43].szAccessPath = L"";
	itemArray[43].szItemID = bstrPH65HS;
	itemArray[43].bActive = true; // active state
	itemArray[43].hClient = (OPCHANDLE)43; // our handle to item
	itemArray[43].dwBlobSize = 0; // no blob support
	itemArray[43].pBlob = NULL;
	itemArray[43].vtRequestedDataType = VT_R4;
	itemArray[43].wReserved = 0;

	/*温度*/
	CString strPH65DT = "PH1_VALUE_PH65_DT";
	BSTR bstrPH65DT = strPH65DT.AllocSysString();
	itemArray[44].szAccessPath = L"";
	itemArray[44].szItemID = bstrPH65DT;
	itemArray[44].bActive = true; // active state
	itemArray[44].hClient = (OPCHANDLE)44; // our handle to item
	itemArray[44].dwBlobSize = 0; // no blob support
	itemArray[44].pBlob = NULL;
	itemArray[44].vtRequestedDataType = VT_R4;
	itemArray[44].wReserved = 0;

	/*压力*/
	CString strPH65DP = "PH1_VALUE_PH65_DP";
	BSTR bstrPH65DP = strPH65DP.AllocSysString();
	itemArray[45].szAccessPath = L"";
	itemArray[45].szItemID = bstrPH65DP;
	itemArray[45].bActive = true; // active state
	itemArray[45].hClient = (OPCHANDLE)45; // our handle to item
	itemArray[45].dwBlobSize = 0; // no blob support
	itemArray[45].pBlob = NULL;
	itemArray[45].vtRequestedDataType = VT_R4;
	itemArray[45].wReserved = 0;

	/*瞬时流量*/
	CString strPH65DF = "PH1_VALUE_PH65_DF";
	BSTR bstrPH65DF = strPH65DF.AllocSysString();
	itemArray[46].szAccessPath = L"";
	itemArray[46].szItemID = bstrPH65DF;
	itemArray[46].bActive = true; // active state
	itemArray[46].hClient = (OPCHANDLE)46; // our handle to item
	itemArray[46].dwBlobSize = 0; // no blob support
	itemArray[46].pBlob = NULL;
	itemArray[46].vtRequestedDataType = VT_R4;
	itemArray[46].wReserved = 0;

	/*累计流量*/
	CString strPH65DS = "PH1_VALUE_PH65_DS";
	BSTR bstrPH65DS = strPH65DS.AllocSysString();
	itemArray[47].szAccessPath = L"";
	itemArray[47].szItemID = bstrPH65DS;
	itemArray[47].bActive = true; // active state
	itemArray[47].hClient = (OPCHANDLE)47; // our handle to item
	itemArray[47].dwBlobSize = 0; // no blob support
	itemArray[47].pBlob = NULL;
	itemArray[47].vtRequestedDataType = VT_R4;
	itemArray[47].wReserved = 0;

	/* PH55_09阀组（9#）*/

	/*汇管温度*/
	CString strPH55HT = "PH1_VALUE_PH55_HT";
	BSTR bstrPH55HT = strPH55HT.AllocSysString();
	itemArray[48].szAccessPath = L"";
	itemArray[48].szItemID = bstrPH55HT;
	itemArray[48].bActive = true; // active state
	itemArray[48].hClient = (OPCHANDLE)48; // our handle to item
	itemArray[48].dwBlobSize = 0; // no blob support
	itemArray[48].pBlob = NULL;
	itemArray[48].vtRequestedDataType = VT_R4;
	itemArray[48].wReserved = 0;

	/*汇管压力*/
	CString strPH55HP = "PH1_VALUE_PH55_HP";
	BSTR bstrPH55HP = strPH55HP.AllocSysString();
	itemArray[49].szAccessPath = L"";
	itemArray[49].szItemID = bstrPH55HP;
	itemArray[49].bActive = true; // active state
	itemArray[49].hClient = (OPCHANDLE)49; // our handle to item
	itemArray[49].dwBlobSize = 0; // no blob support
	itemArray[49].pBlob = NULL;
	itemArray[49].vtRequestedDataType = VT_R4;
	itemArray[49].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH55HF = "PH1_VALUE_PH55_HF";
	BSTR bstrPH55HF = strPH55HF.AllocSysString();
	itemArray[50].szAccessPath = L"";
	itemArray[50].szItemID = bstrPH55HF;
	itemArray[50].bActive = true; // active state
	itemArray[50].hClient = (OPCHANDLE)50; // our handle to item
	itemArray[50].dwBlobSize = 0; // no blob support
	itemArray[50].pBlob = NULL;
	itemArray[50].vtRequestedDataType = VT_R4;
	itemArray[50].wReserved = 0;

	/*汇管累计流量*/
	CString strPH55HS = "PH1_VALUE_PH55_HS";
	BSTR bstrPH55HS = strPH55HS.AllocSysString();
	itemArray[51].szAccessPath = L"";
	itemArray[51].szItemID = bstrPH55HS;
	itemArray[51].bActive = true; // active state
	itemArray[51].hClient = (OPCHANDLE)51; // our handle to item
	itemArray[51].dwBlobSize = 0; // no blob support
	itemArray[51].pBlob = NULL;
	itemArray[51].vtRequestedDataType = VT_R4;
	itemArray[51].wReserved = 0;

	/*温度*/
	CString strPH55DT = "PH1_VALUE_PH55_DT";
	BSTR bstrPH55DT = strPH55DT.AllocSysString();
	itemArray[52].szAccessPath = L"";
	itemArray[52].szItemID = bstrPH55DT;
	itemArray[52].bActive = true; // active state
	itemArray[52].hClient = (OPCHANDLE)52; // our handle to item
	itemArray[52].dwBlobSize = 0; // no blob support
	itemArray[52].pBlob = NULL;
	itemArray[52].vtRequestedDataType = VT_R4;
	itemArray[52].wReserved = 0;

	/*压力*/
	CString strPH55DP = "PH1_VALUE_PH55_DP";
	BSTR bstrPH55DP = strPH55DP.AllocSysString();
	itemArray[53].szAccessPath = L"";
	itemArray[53].szItemID = bstrPH55DP;
	itemArray[53].bActive = true; // active state
	itemArray[53].hClient = (OPCHANDLE)53; // our handle to item
	itemArray[53].dwBlobSize = 0; // no blob support
	itemArray[53].pBlob = NULL;
	itemArray[53].vtRequestedDataType = VT_R4;
	itemArray[53].wReserved = 0;

	/*瞬时流量*/
	CString strPH55DF = "PH1_VALUE_PH55_DF";
	BSTR bstrPH55DF = strPH55DF.AllocSysString();
	itemArray[54].szAccessPath = L"";
	itemArray[54].szItemID = bstrPH55DF;
	itemArray[54].bActive = true; // active state
	itemArray[54].hClient = (OPCHANDLE)54; // our handle to item
	itemArray[54].dwBlobSize = 0; // no blob support
	itemArray[54].pBlob = NULL;
	itemArray[54].vtRequestedDataType = VT_R4;
	itemArray[54].wReserved = 0;

	/*累计流量*/
	CString strPH55DS = "PH1_VALUE_PH55_DS";
	BSTR bstrPH55DS = strPH55DS.AllocSysString();
	itemArray[55].szAccessPath = L"";
	itemArray[55].szItemID = bstrPH55DS;
	itemArray[55].bActive = true; // active state
	itemArray[55].hClient = (OPCHANDLE)55; // our handle to item
	itemArray[55].dwBlobSize = 0; // no blob support
	itemArray[55].pBlob = NULL;
	itemArray[55].vtRequestedDataType = VT_R4;
	itemArray[55].wReserved = 0;

	/* PH76_03阀组（10#）*/

	/*汇管温度*/
	CString strPH76HT = "PH1_VALUE_PH76_HT";
	BSTR bstrPH76HT = strPH76HT.AllocSysString();
	itemArray[56].szAccessPath = L"";
	itemArray[56].szItemID = bstrPH76HT;
	itemArray[56].bActive = true; // active state
	itemArray[56].hClient = (OPCHANDLE)56; // our handle to item
	itemArray[56].dwBlobSize = 0; // no blob support
	itemArray[56].pBlob = NULL;
	itemArray[56].vtRequestedDataType = VT_R4;
	itemArray[56].wReserved = 0;

	/*汇管压力*/
	CString strPH76HP = "PH1_VALUE_PH76_HP";
	BSTR bstrPH76HP = strPH76HP.AllocSysString();
	itemArray[57].szAccessPath = L"";
	itemArray[57].szItemID = bstrPH76HP;
	itemArray[57].bActive = true; // active state
	itemArray[57].hClient = (OPCHANDLE)57; // our handle to item
	itemArray[57].dwBlobSize = 0; // no blob support
	itemArray[57].pBlob = NULL;
	itemArray[57].vtRequestedDataType = VT_R4;
	itemArray[57].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH76HF = "PH1_VALUE_PH76_HF";
	BSTR bstrPH76HF = strPH76HF.AllocSysString();
	itemArray[58].szAccessPath = L"";
	itemArray[58].szItemID = bstrPH76HF;
	itemArray[58].bActive = true; // active state
	itemArray[58].hClient = (OPCHANDLE)58; // our handle to item
	itemArray[58].dwBlobSize = 0; // no blob support
	itemArray[58].pBlob = NULL;
	itemArray[58].vtRequestedDataType = VT_R4;
	itemArray[58].wReserved = 0;

	/*汇管累计流量*/
	CString strPH76HS = "PH1_VALUE_PH76_HS";
	BSTR bstrPH76HS = strPH76HS.AllocSysString();
	itemArray[59].szAccessPath = L"";
	itemArray[59].szItemID = bstrPH76HS;
	itemArray[59].bActive = true; // active state
	itemArray[59].hClient = (OPCHANDLE)59; // our handle to item
	itemArray[59].dwBlobSize = 0; // no blob support
	itemArray[59].pBlob = NULL;
	itemArray[59].vtRequestedDataType = VT_R4;
	itemArray[59].wReserved = 0;

	/*温度*/
	CString strPH76DT = "PH1_VALUE_PH76_DT";
	BSTR bstrPH76DT = strPH76DT.AllocSysString();
	itemArray[60].szAccessPath = L"";
	itemArray[60].szItemID = bstrPH76DT;
	itemArray[60].bActive = true; // active state
	itemArray[60].hClient = (OPCHANDLE)60; // our handle to item
	itemArray[60].dwBlobSize = 0; // no blob support
	itemArray[60].pBlob = NULL;
	itemArray[60].vtRequestedDataType = VT_R4;
	itemArray[60].wReserved = 0;

	/*压力*/
	CString strPH76DP = "PH1_VALUE_PH76_DP";
	BSTR bstrPH76DP = strPH76DP.AllocSysString();
	itemArray[61].szAccessPath = L"";
	itemArray[61].szItemID = bstrPH76DP;
	itemArray[61].bActive = true; // active state
	itemArray[61].hClient = (OPCHANDLE)61; // our handle to item
	itemArray[61].dwBlobSize = 0; // no blob support
	itemArray[61].pBlob = NULL;
	itemArray[61].vtRequestedDataType = VT_R4;
	itemArray[61].wReserved = 0;

	/*瞬时流量*/
	CString strPH76DF = "PH1_VALUE_PH76_DF";
	BSTR bstrPH76DF = strPH76DF.AllocSysString();
	itemArray[62].szAccessPath = L"";
	itemArray[62].szItemID = bstrPH76DF;
	itemArray[62].bActive = true; // active state
	itemArray[62].hClient = (OPCHANDLE)62; // our handle to item
	itemArray[62].dwBlobSize = 0; // no blob support
	itemArray[62].pBlob = NULL;
	itemArray[62].vtRequestedDataType = VT_R4;
	itemArray[62].wReserved = 0;

	/*累计流量*/
	CString strPH76DS = "PH1_VALUE_PH76_DS";
	BSTR bstrPH76DS = strPH76DS.AllocSysString();
	itemArray[63].szAccessPath = L"";
	itemArray[63].szItemID = bstrPH76DS;
	itemArray[63].bActive = true; // active state
	itemArray[63].hClient = (OPCHANDLE)63; // our handle to item
	itemArray[63].dwBlobSize = 23; // no blob support
	itemArray[63].pBlob = NULL;
	itemArray[63].vtRequestedDataType = VT_R4;
	itemArray[63].wReserved = 0;

	/* PH66_01阀组（11#）*/

	/*汇管温度*/
	CString strPH66HT = "PH1_VALUE_PH66_HT";
	BSTR bstrPH66HT = strPH66HT.AllocSysString();
	itemArray[64].szAccessPath = L"";
	itemArray[64].szItemID = bstrPH66HT;
	itemArray[64].bActive = true; // active state
	itemArray[64].hClient = (OPCHANDLE)64; // our handle to item
	itemArray[64].dwBlobSize = 0; // no blob support
	itemArray[64].pBlob = NULL;
	itemArray[64].vtRequestedDataType = VT_R4;
	itemArray[64].wReserved = 0;

	/*汇管压力*/
	CString strPH66HP = "PH1_VALUE_PH66_HP";
	BSTR bstrPH66HP = strPH66HP.AllocSysString();
	itemArray[65].szAccessPath = L"";
	itemArray[65].szItemID = bstrPH66HP;
	itemArray[65].bActive = true; // active state
	itemArray[65].hClient = (OPCHANDLE)65; // our handle to item
	itemArray[65].dwBlobSize = 0; // no blob support
	itemArray[65].pBlob = NULL;
	itemArray[65].vtRequestedDataType = VT_R4;
	itemArray[65].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH66HF = "PH1_VALUE_PH66_HF";
	BSTR bstrPH66HF = strPH66HF.AllocSysString();
	itemArray[66].szAccessPath = L"";
	itemArray[66].szItemID = bstrPH66HF;
	itemArray[66].bActive = true; // active state
	itemArray[66].hClient = (OPCHANDLE)66; // our handle to item
	itemArray[66].dwBlobSize = 0; // no blob support
	itemArray[66].pBlob = NULL;
	itemArray[66].vtRequestedDataType = VT_R4;
	itemArray[66].wReserved = 0;

	/*汇管累计流量*/
	CString strPH66HS = "PH1_VALUE_PH66_HS";
	BSTR bstrPH66HS = strPH66HS.AllocSysString();
	itemArray[67].szAccessPath = L"";
	itemArray[67].szItemID = bstrPH66HS;
	itemArray[67].bActive = true; // active state
	itemArray[67].hClient = (OPCHANDLE)67; // our handle to item
	itemArray[67].dwBlobSize = 0; // no blob support
	itemArray[67].pBlob = NULL;
	itemArray[67].vtRequestedDataType = VT_R4;
	itemArray[67].wReserved = 0;

	/*温度*/
	CString strPH66DT = "PH1_VALUE_PH66_DT";
	BSTR bstrPH66DT = strPH66DT.AllocSysString();
	itemArray[68].szAccessPath = L"";
	itemArray[68].szItemID = bstrPH66DT;
	itemArray[68].bActive = true; // active state
	itemArray[68].hClient = (OPCHANDLE)68; // our handle to item
	itemArray[68].dwBlobSize = 0; // no blob support
	itemArray[68].pBlob = NULL;
	itemArray[68].vtRequestedDataType = VT_R4;
	itemArray[68].wReserved = 0;

	/*压力*/
	CString strPH66DP = "PH1_VALUE_PH66_DP";
	BSTR bstrPH66DP = strPH66DP.AllocSysString();
	itemArray[69].szAccessPath = L"";
	itemArray[69].szItemID = bstrPH66DP;
	itemArray[69].bActive = true; // active state
	itemArray[69].hClient = (OPCHANDLE)69; // our handle to item
	itemArray[69].dwBlobSize = 0; // no blob support
	itemArray[69].pBlob = NULL;
	itemArray[69].vtRequestedDataType = VT_R4;
	itemArray[69].wReserved = 0;

	/*瞬时流量*/
	CString strPH66DF = "PH1_VALUE_PH66_DF";
	BSTR bstrPH66DF = strPH66DF.AllocSysString();
	itemArray[70].szAccessPath = L"";
	itemArray[70].szItemID = bstrPH66DF;
	itemArray[70].bActive = true; // active state
	itemArray[70].hClient = (OPCHANDLE)70; // our handle to item
	itemArray[70].dwBlobSize = 0; // no blob support
	itemArray[70].pBlob = NULL;
	itemArray[70].vtRequestedDataType = VT_R4;
	itemArray[70].wReserved = 0;

	/*累计流量*/
	CString strPH66DS = "PH1_VALUE_PH66_DS";
	BSTR bstrPH66DS = strPH66DS.AllocSysString();
	itemArray[71].szAccessPath = L"";
	itemArray[71].szItemID = bstrPH66DS;
	itemArray[71].bActive = true; // active state
	itemArray[71].hClient = (OPCHANDLE)71; // our handle to item
	itemArray[71].dwBlobSize = 0; // no blob support
	itemArray[71].pBlob = NULL;
	itemArray[71].vtRequestedDataType = VT_R4;
	itemArray[71].wReserved = 0;

	/* PH77_08阀组（12#）*/

	/*汇管温度*/
	CString strPH77HT = "PH1_VALUE_PH77_HT";
	BSTR bstrPH77HT = strPH77HT.AllocSysString();
	itemArray[72].szAccessPath = L"";
	itemArray[72].szItemID = bstrPH77HT;
	itemArray[72].bActive = true; // active state
	itemArray[72].hClient = (OPCHANDLE)72; // our handle to item
	itemArray[72].dwBlobSize = 0; // no blob support
	itemArray[72].pBlob = NULL;
	itemArray[72].vtRequestedDataType = VT_R4;
	itemArray[72].wReserved = 0;

	/*汇管压力*/
	CString strPH77HP = "PH1_VALUE_PH77_HP";
	BSTR bstrPH77HP = strPH77HP.AllocSysString();
	itemArray[73].szAccessPath = L"";
	itemArray[73].szItemID = bstrPH77HP;
	itemArray[73].bActive = true; // active state
	itemArray[73].hClient = (OPCHANDLE)73; // our handle to item
	itemArray[73].dwBlobSize = 0; // no blob support
	itemArray[73].pBlob = NULL;
	itemArray[73].vtRequestedDataType = VT_R4;
	itemArray[73].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH77HF = "PH1_VALUE_PH77_HF";
	BSTR bstrPH77HF = strPH77HF.AllocSysString();
	itemArray[74].szAccessPath = L"";
	itemArray[74].szItemID = bstrPH77HF;
	itemArray[74].bActive = true; // active state
	itemArray[74].hClient = (OPCHANDLE)74; // our handle to item
	itemArray[74].dwBlobSize = 0; // no blob support
	itemArray[74].pBlob = NULL;
	itemArray[74].vtRequestedDataType = VT_R4;
	itemArray[74].wReserved = 0;

	/*汇管累计流量*/
	CString strPH77HS = "PH1_VALUE_PH77_HS";
	BSTR bstrPH77HS = strPH77HS.AllocSysString();
	itemArray[75].szAccessPath = L"";
	itemArray[75].szItemID = bstrPH77HS;
	itemArray[75].bActive = true; // active state
	itemArray[75].hClient = (OPCHANDLE)75; // our handle to item
	itemArray[75].dwBlobSize = 0; // no blob support
	itemArray[75].pBlob = NULL;
	itemArray[75].vtRequestedDataType = VT_R4;
	itemArray[75].wReserved = 0;

	/*温度*/
	CString strPH77DT = "PH1_VALUE_PH77_DT";
	BSTR bstrPH77DT = strPH77DT.AllocSysString();
	itemArray[76].szAccessPath = L"";
	itemArray[76].szItemID = bstrPH77DT;
	itemArray[76].bActive = true; // active state
	itemArray[76].hClient = (OPCHANDLE)76; // our handle to item
	itemArray[76].dwBlobSize = 0; // no blob support
	itemArray[76].pBlob = NULL;
	itemArray[76].vtRequestedDataType = VT_R4;
	itemArray[76].wReserved = 0;

	/*压力*/
	CString strPH77DP = "PH1_VALUE_PH77_DP";
	BSTR bstrPH77DP = strPH77DP.AllocSysString();
	itemArray[77].szAccessPath = L"";
	itemArray[77].szItemID = bstrPH77DP;
	itemArray[77].bActive = true; // active state
	itemArray[77].hClient = (OPCHANDLE)77; // our handle to item
	itemArray[77].dwBlobSize = 0; // no blob support
	itemArray[77].pBlob = NULL;
	itemArray[77].vtRequestedDataType = VT_R4;
	itemArray[77].wReserved = 0;

	/*瞬时流量*/
	CString strPH77DF = "PH1_VALUE_PH77_DF";
	BSTR bstrPH77DF = strPH77DF.AllocSysString();
	itemArray[78].szAccessPath = L"";
	itemArray[78].szItemID = bstrPH77DF;
	itemArray[78].bActive = true; // active state
	itemArray[78].hClient = (OPCHANDLE)78; // our handle to item
	itemArray[78].dwBlobSize = 0; // no blob support
	itemArray[78].pBlob = NULL;
	itemArray[78].vtRequestedDataType = VT_R4;
	itemArray[78].wReserved = 0;

	/*累计流量*/
	CString strPH77DS = "PH1_VALUE_PH77_DS";
	BSTR bstrPH77DS = strPH77DS.AllocSysString();
	itemArray[79].szAccessPath = L"";
	itemArray[79].szItemID = bstrPH77DS;
	itemArray[79].bActive = true; // active state
	itemArray[79].hClient = (OPCHANDLE)79; // our handle to item
	itemArray[79].dwBlobSize = 0; // no blob support
	itemArray[79].pBlob = NULL;
	itemArray[79].vtRequestedDataType = VT_R4;
	itemArray[79].wReserved = 0;

	/* PE_079阀组（14#）*/

	/*汇管温度*/
	CString strPE79HT = "PH1_VALUE_PE79_HT";
	BSTR bstrPE79HT = strPE79HT.AllocSysString();
	itemArray[80].szAccessPath = L"";
	itemArray[80].szItemID = bstrPE79HT;
	itemArray[80].bActive = true; // active state
	itemArray[80].hClient = (OPCHANDLE)80; // our handle to item
	itemArray[80].dwBlobSize = 0; // no blob support
	itemArray[80].pBlob = NULL;
	itemArray[80].vtRequestedDataType = VT_R4;
	itemArray[80].wReserved = 0;

	/*汇管压力*/
	CString strPE79HP = "PH1_VALUE_PE79_HP";
	BSTR bstrPE79HP = strPE79HP.AllocSysString();
	itemArray[81].szAccessPath = L"";
	itemArray[81].szItemID = bstrPE79HP;
	itemArray[81].bActive = true; // active state
	itemArray[81].hClient = (OPCHANDLE)81; // our handle to item
	itemArray[81].dwBlobSize = 0; // no blob support
	itemArray[81].pBlob = NULL;
	itemArray[81].vtRequestedDataType = VT_R4;
	itemArray[81].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPE79HF = "PH1_VALUE_PE79_HF";
	BSTR bstrPE79HF = strPE79HF.AllocSysString();
	itemArray[82].szAccessPath = L"";
	itemArray[82].szItemID = bstrPE79HF;
	itemArray[82].bActive = true; // active state
	itemArray[82].hClient = (OPCHANDLE)82; // our handle to item
	itemArray[82].dwBlobSize = 0; // no blob support
	itemArray[82].pBlob = NULL;
	itemArray[82].vtRequestedDataType = VT_R4;
	itemArray[82].wReserved = 0;

	/*汇管累计流量*/
	CString strPE79HS = "PH1_VALUE_PE79_HS";
	BSTR bstrPE79HS = strPE79HS.AllocSysString();
	itemArray[83].szAccessPath = L"";
	itemArray[83].szItemID = bstrPE79HS;
	itemArray[83].bActive = true; // active state
	itemArray[83].hClient = (OPCHANDLE)83; // our handle to item
	itemArray[83].dwBlobSize = 0; // no blob support
	itemArray[83].pBlob = NULL;
	itemArray[83].vtRequestedDataType = VT_R4;
	itemArray[83].wReserved = 0;

	/*温度*/
	CString strPE79DT = "PH1_VALUE_PE79_DT";
	BSTR bstrPE79DT = strPE79DT.AllocSysString();
	itemArray[84].szAccessPath = L"";
	itemArray[84].szItemID = bstrPE79DT;
	itemArray[84].bActive = true; // active state
	itemArray[84].hClient = (OPCHANDLE)84; // our handle to item
	itemArray[84].dwBlobSize = 0; // no blob support
	itemArray[84].pBlob = NULL;
	itemArray[84].vtRequestedDataType = VT_R4;
	itemArray[84].wReserved = 0;

	/*压力*/
	CString strPE79DP = "PH1_VALUE_PE79_DP";
	BSTR bstrPE79DP = strPE79DP.AllocSysString();
	itemArray[85].szAccessPath = L"";
	itemArray[85].szItemID = bstrPE79DP;
	itemArray[85].bActive = true; // active state
	itemArray[85].hClient = (OPCHANDLE)85; // our handle to item
	itemArray[85].dwBlobSize = 0; // no blob support
	itemArray[85].pBlob = NULL;
	itemArray[85].vtRequestedDataType = VT_R4;
	itemArray[85].wReserved = 0;

	/*瞬时流量*/
	CString strPE79DF = "PH1_VALUE_PE79_DF";
	BSTR bstrPE79DF = strPE79DF.AllocSysString();
	itemArray[86].szAccessPath = L"";
	itemArray[86].szItemID = bstrPE79DF;
	itemArray[86].bActive = true; // active state
	itemArray[86].hClient = (OPCHANDLE)86; // our handle to item
	itemArray[86].dwBlobSize = 0; // no blob support
	itemArray[86].pBlob = NULL;
	itemArray[86].vtRequestedDataType = VT_R4;
	itemArray[86].wReserved = 0;

	/*累计流量*/
	CString strPE79DS = "PH1_VALUE_PE79_DS";
	BSTR bstrPE79DS = strPE79DS.AllocSysString();
	itemArray[87].szAccessPath = L"";
	itemArray[87].szItemID = bstrPE79DS;
	itemArray[87].bActive = true; // active state
	itemArray[87].hClient = (OPCHANDLE)87; // our handle to item
	itemArray[87].dwBlobSize = 0; // no blob support
	itemArray[87].pBlob = NULL;
	itemArray[87].vtRequestedDataType = VT_R4;
	itemArray[87].wReserved = 0;

	/* PE_112阀组（15#）*/

	/*汇管温度*/
	CString strPE112HT = "PH1_VALUE_PE112_HT";
	BSTR bstrPE112HT = strPE112HT.AllocSysString();
	itemArray[88].szAccessPath = L"";
	itemArray[88].szItemID = bstrPE112HT;
	itemArray[88].bActive = true; // active state
	itemArray[88].hClient = (OPCHANDLE)88; // our handle to item
	itemArray[88].dwBlobSize = 0; // no blob support
	itemArray[88].pBlob = NULL;
	itemArray[88].vtRequestedDataType = VT_R4;
	itemArray[88].wReserved = 0;

	/*汇管压力*/
	CString strPE112HP = "PH1_VALUE_PE112_HP";
	BSTR bstrPE112HP = strPE112HP.AllocSysString();
	itemArray[89].szAccessPath = L"";
	itemArray[89].szItemID = bstrPE112HP;
	itemArray[89].bActive = true; // active state
	itemArray[89].hClient = (OPCHANDLE)89; // our handle to item
	itemArray[89].dwBlobSize = 0; // no blob support
	itemArray[89].pBlob = NULL;
	itemArray[89].vtRequestedDataType = VT_R4;
	itemArray[89].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPE112HF = "PH1_VALUE_PE112_HF";
	BSTR bstrPE112HF = strPE112HF.AllocSysString();
	itemArray[90].szAccessPath = L"";
	itemArray[90].szItemID = bstrPE112HF;
	itemArray[90].bActive = true; // active state
	itemArray[90].hClient = (OPCHANDLE)90; // our handle to item
	itemArray[90].dwBlobSize = 0; // no blob support
	itemArray[90].pBlob = NULL;
	itemArray[90].vtRequestedDataType = VT_R4;
	itemArray[90].wReserved = 0;

	/*汇管累计流量*/
	CString strPE112HS = "PH1_VALUE_PE112_HS";
	BSTR bstrPE112HS = strPE112HS.AllocSysString();
	itemArray[91].szAccessPath = L"";
	itemArray[91].szItemID = bstrPE112HS;
	itemArray[91].bActive = true; // active state
	itemArray[91].hClient = (OPCHANDLE)91; // our handle to item
	itemArray[91].dwBlobSize = 0; // no blob support
	itemArray[91].pBlob = NULL;
	itemArray[91].vtRequestedDataType = VT_R4;
	itemArray[91].wReserved = 0;

	/*温度*/
	CString strPE112DT = "PH1_VALUE_PE112_DT";
	BSTR bstrPE112DT = strPE112DT.AllocSysString();
	itemArray[92].szAccessPath = L"";
	itemArray[92].szItemID = bstrPE112DT;
	itemArray[92].bActive = true; // active state
	itemArray[92].hClient = (OPCHANDLE)92; // our handle to item
	itemArray[92].dwBlobSize = 0; // no blob support
	itemArray[92].pBlob = NULL;
	itemArray[92].vtRequestedDataType = VT_R4;
	itemArray[92].wReserved = 0;

	/*压力*/
	CString strPE112DP = "PH1_VALUE_PE112_DP";
	BSTR bstrPE112DP = strPE112DP.AllocSysString();
	itemArray[93].szAccessPath = L"";
	itemArray[93].szItemID = bstrPE112DP;
	itemArray[93].bActive = true; // active state
	itemArray[93].hClient = (OPCHANDLE)93; // our handle to item
	itemArray[93].dwBlobSize = 0; // no blob support
	itemArray[93].pBlob = NULL;
	itemArray[93].vtRequestedDataType = VT_R4;
	itemArray[93].wReserved = 0;

	/*瞬时流量*/
	CString strPE112DF = "PH1_VALUE_PE112_DF";
	BSTR bstrPE112DF = strPE112DF.AllocSysString();
	itemArray[94].szAccessPath = L"";
	itemArray[94].szItemID = bstrPE112DF;
	itemArray[94].bActive = true; // active state
	itemArray[94].hClient = (OPCHANDLE)94; // our handle to item
	itemArray[94].dwBlobSize = 0; // no blob support
	itemArray[94].pBlob = NULL;
	itemArray[94].vtRequestedDataType = VT_R4;
	itemArray[94].wReserved = 0;

	/*累计流量*/
	CString strPE112DS = "PH1_VALUE_PE112_DS";
	BSTR bstrPE112DS = strPE112DS.AllocSysString();
	itemArray[95].szAccessPath = L"";
	itemArray[95].szItemID = bstrPE112DS;
	itemArray[95].bActive = true; // active state
	itemArray[95].hClient = (OPCHANDLE)95; // our handle to item
	itemArray[95].dwBlobSize = 0; // no blob support
	itemArray[95].pBlob = NULL;
	itemArray[95].vtRequestedDataType = VT_R4;
	itemArray[95].wReserved = 0;

	/* PH73-11阀组（3#）*/

	/*汇管温度*/
	CString strPH73HT = "PH1_VALUE_PH73_HT";
	BSTR bstrPH73HT = strPH73HT.AllocSysString();
	itemArray[96].szAccessPath = L"";
	itemArray[96].szItemID = bstrPH73HT;
	itemArray[96].bActive = true; // active state
	itemArray[96].hClient = (OPCHANDLE)96; // our handle to item
	itemArray[96].dwBlobSize = 0; // no blob support
	itemArray[96].pBlob = NULL;
	itemArray[96].vtRequestedDataType = VT_R4;
	itemArray[96].wReserved = 0;

	/*汇管压力*/
	CString strPH73HP = "PH1_VALUE_PH73_HP";
	BSTR bstrPH73HP = strPH73HP.AllocSysString();
	itemArray[97].szAccessPath = L"";
	itemArray[97].szItemID = bstrPH73HP;
	itemArray[97].bActive = true; // active state
	itemArray[97].hClient = (OPCHANDLE)97; // our handle to item
	itemArray[97].dwBlobSize = 0; // no blob support
	itemArray[97].pBlob = NULL;
	itemArray[97].vtRequestedDataType = VT_R4;
	itemArray[97].wReserved = 0;

	/*汇管瞬时流量*/
	CString strPH73HF = "PH1_VALUE_PH73_HF";
	BSTR bstrPH73HF = strPH73HF.AllocSysString();
	itemArray[98].szAccessPath = L"";
	itemArray[98].szItemID = bstrPH73HF;
	itemArray[98].bActive = true; // active state
	itemArray[98].hClient = (OPCHANDLE)98; // our handle to item
	itemArray[98].dwBlobSize = 0; // no blob support
	itemArray[98].pBlob = NULL;
	itemArray[98].vtRequestedDataType = VT_R4;
	itemArray[98].wReserved = 0;

	/*汇管累计流量*/
	CString strPH73HS = "PH1_VALUE_PH73_HS";
	BSTR bstrPH73HS = strPH73HS.AllocSysString();
	itemArray[99].szAccessPath = L"";
	itemArray[99].szItemID = bstrPH73HS;
	itemArray[99].bActive = true; // active state
	itemArray[99].hClient = (OPCHANDLE)99; // our handle to item
	itemArray[99].dwBlobSize = 0; // no blob support
	itemArray[99].pBlob = NULL;
	itemArray[99].vtRequestedDataType = VT_R4;
	itemArray[99].wReserved = 0;

	/*温度*/
	CString strPH73DT = "PH1_VALUE_PH73_DT";
	BSTR bstrPH73DT = strPH73DT.AllocSysString();
	itemArray[100].szAccessPath = L"";
	itemArray[100].szItemID = bstrPH73DT;
	itemArray[100].bActive = true; // active state
	itemArray[100].hClient = (OPCHANDLE)100; // our handle to item
	itemArray[100].dwBlobSize = 0; // no blob support
	itemArray[100].pBlob = NULL;
	itemArray[100].vtRequestedDataType = VT_R4;
	itemArray[100].wReserved = 0;

	/*压力*/
	CString strPH73DP = "PH1_VALUE_PH73_DP";
	BSTR bstrPH73DP = strPH73DP.AllocSysString();
	itemArray[101].szAccessPath = L"";
	itemArray[101].szItemID = bstrPH73DP;
	itemArray[101].bActive = true; // active state
	itemArray[101].hClient = (OPCHANDLE)101; // our handle to item
	itemArray[101].dwBlobSize = 0; // no blob support
	itemArray[101].pBlob = NULL;
	itemArray[101].vtRequestedDataType = VT_R4;
	itemArray[101].wReserved = 0;

	/*瞬时流量*/
	CString strPH73DF = "PH1_VALUE_PH73_DF";
	BSTR bstrPH73DF = strPH73DF.AllocSysString();
	itemArray[102].szAccessPath = L"";
	itemArray[102].szItemID = bstrPH73DF;
	itemArray[102].bActive = true; // active state
	itemArray[102].hClient = (OPCHANDLE)102; // our handle to item
	itemArray[102].dwBlobSize = 0; // no blob support
	itemArray[102].pBlob = NULL;
	itemArray[102].vtRequestedDataType = VT_R4;
	itemArray[102].wReserved = 0;

	/*累计流量*/
	CString strPH73DS = "PH1_VALUE_PH73_DS";
	BSTR bstrPH73DS = strPH73DS.AllocSysString();
	itemArray[103].szAccessPath = L"";
	itemArray[103].szItemID = bstrPH73DS;
	itemArray[103].bActive = true; // active state
	itemArray[103].hClient = (OPCHANDLE)103; // our handle to item
	itemArray[103].dwBlobSize = 23; // no blob support
	itemArray[103].pBlob = NULL;
	itemArray[103].vtRequestedDataType = VT_R4;
	itemArray[103].wReserved = 0;

	/* PH45-05单井轮换*/

	/* 汇管压力*/
	CString strPH45HP = "PH1_VALUE_PH45_HP";
	BSTR bstrPH45HP = strPH45HP.AllocSysString();
	itemArray[104].szAccessPath = L"";
	itemArray[104].szItemID = bstrPH45HP;
	itemArray[104].bActive = true; // active state
	itemArray[104].hClient = (OPCHANDLE)104; // our handle to item
	itemArray[104].dwBlobSize = 0; // no blob support
	itemArray[104].pBlob = NULL;
	itemArray[104].vtRequestedDataType = VT_R4;
	itemArray[104].wReserved = 0;

	/* 单井温度*/
	CString strPH45DT = "PH1_VALUE_PH45_DT";
	BSTR bstrPH45DT = strPH45DT.AllocSysString();
	itemArray[105].szAccessPath = L"";
	itemArray[105].szItemID = bstrPH45DT;
	itemArray[105].bActive = true; // active state
	itemArray[105].hClient = (OPCHANDLE)105; // our handle to item
	itemArray[105].dwBlobSize = 0; // no blob support
	itemArray[105].pBlob = NULL;
	itemArray[105].vtRequestedDataType = VT_R4;
	itemArray[105].wReserved = 0;

	/* 单井压力*/
	CString strPH45DP = "PH1_VALUE_PH45_DP";
	BSTR bstrPH45DP = strPH45DP.AllocSysString();
	itemArray[106].szAccessPath = L"";
	itemArray[106].szItemID = bstrPH45DP;
	itemArray[106].bActive = true; // active state
	itemArray[106].hClient = (OPCHANDLE)106; // our handle to item
	itemArray[106].dwBlobSize = 0; // no blob support
	itemArray[106].pBlob = NULL;
	itemArray[106].vtRequestedDataType = VT_R4;
	itemArray[106].wReserved = 0;

	/*瞬时流量*/
	CString strPH45F = "PH1_VALUE_PH45_F";
	BSTR bstrPH45F = strPH45F.AllocSysString();
	itemArray[107].szAccessPath = L"";
	itemArray[107].szItemID = bstrPH45F;
	itemArray[107].bActive = true; // active state
	itemArray[107].hClient = (OPCHANDLE)107; // our handle to item
	itemArray[107].dwBlobSize = 0; // no blob support
	itemArray[107].pBlob = NULL;
	itemArray[107].vtRequestedDataType = VT_R4;
	itemArray[107].wReserved = 0;

	/*累计流量*/
	CString strPH45S = "PH1_VALUE_PH45_S";
	BSTR bstrPH45S = strPH45S.AllocSysString();
	itemArray[108].szAccessPath = L"";
	itemArray[108].szItemID = bstrPH45S;
	itemArray[108].bActive = true; // active state
	itemArray[108].hClient = (OPCHANDLE)108; // our handle to item
	itemArray[108].dwBlobSize = 23; // no blob support
	itemArray[108].pBlob = NULL;
	itemArray[108].vtRequestedDataType = VT_R4;
	itemArray[108].wReserved = 0;

	/* PH35-07单井轮换*/

	/* 汇管压力*/
	CString strPH35HP = "PH1_VALUE_PH35_HP";
	BSTR bstrPH35HP = strPH35HP.AllocSysString();
	itemArray[109].szAccessPath = L"";
	itemArray[109].szItemID = bstrPH35HP;
	itemArray[109].bActive = true; // active state
	itemArray[109].hClient = (OPCHANDLE)109; // our handle to item
	itemArray[109].dwBlobSize = 0; // no blob support
	itemArray[109].pBlob = NULL;
	itemArray[109].vtRequestedDataType = VT_R4;
	itemArray[109].wReserved = 0;

	/* 单井温度*/
	CString strPH35DT = "PH1_VALUE_PH35_DT";
	BSTR bstrPH35DT = strPH35DT.AllocSysString();
	itemArray[110].szAccessPath = L"";
	itemArray[110].szItemID = bstrPH35DT;
	itemArray[110].bActive = true; // active state
	itemArray[110].hClient = (OPCHANDLE)110; // our handle to item
	itemArray[110].dwBlobSize = 0; // no blob support
	itemArray[110].pBlob = NULL;
	itemArray[110].vtRequestedDataType = VT_R4;
	itemArray[110].wReserved = 0;

	/* 单井压力*/
	CString strPH35DP = "PH1_VALUE_PH35_DP";
	BSTR bstrPH35DP = strPH35DP.AllocSysString();
	itemArray[111].szAccessPath = L"";
	itemArray[111].szItemID = bstrPH35DP;
	itemArray[111].bActive = true; // active state
	itemArray[111].hClient = (OPCHANDLE)111; // our handle to item
	itemArray[111].dwBlobSize = 0; // no blob support
	itemArray[111].pBlob = NULL;
	itemArray[111].vtRequestedDataType = VT_R4;
	itemArray[111].wReserved = 0;

	/*瞬时流量*/
	CString strPH35F = "PH1_VALUE_PH35_F";
	BSTR bstrPH35F = strPH35F.AllocSysString();
	itemArray[112].szAccessPath = L"";
	itemArray[112].szItemID = bstrPH35F;
	itemArray[112].bActive = true; // active state
	itemArray[112].hClient = (OPCHANDLE)112; // our handle to item
	itemArray[112].dwBlobSize = 0; // no blob support
	itemArray[112].pBlob = NULL;
	itemArray[112].vtRequestedDataType = VT_R4;
	itemArray[112].wReserved = 0;

	/*累计流量*/
	CString strPH35S = "PH1_VALUE_PH35_S";
	BSTR bstrPH35S = strPH35S.AllocSysString();
	itemArray[113].szAccessPath = L"";
	itemArray[113].szItemID = bstrPH35S;
	itemArray[113].bActive = true; // active state
	itemArray[113].hClient = (OPCHANDLE)113; // our handle to item
	itemArray[113].dwBlobSize = 23; // no blob support
	itemArray[113].pBlob = NULL;
	itemArray[113].vtRequestedDataType = VT_R4;
	itemArray[113].wReserved = 0;

	/* PH1-02单井轮换*/

	/* 汇管压力*/
	CString strPH1HP = "PH1_VALUE_PH1_HP";
	BSTR bstrPH1HP = strPH1HP.AllocSysString();
	itemArray[114].szAccessPath = L"";
	itemArray[114].szItemID = bstrPH1HP;
	itemArray[114].bActive = true; // active state
	itemArray[114].hClient = (OPCHANDLE)114; // our handle to item
	itemArray[114].dwBlobSize = 0; // no blob support
	itemArray[114].pBlob = NULL;
	itemArray[114].vtRequestedDataType = VT_R4;
	itemArray[114].wReserved = 0;

	/* 单井温度*/
	CString strPH1DT = "PH1_VALUE_PH1_DT";
	BSTR bstrPH1DT = strPH1DT.AllocSysString();
	itemArray[115].szAccessPath = L"";
	itemArray[115].szItemID = bstrPH1DT;
	itemArray[115].bActive = true; // active state
	itemArray[115].hClient = (OPCHANDLE)115; // our handle to item
	itemArray[115].dwBlobSize = 0; // no blob support
	itemArray[115].pBlob = NULL;
	itemArray[115].vtRequestedDataType = VT_R4;
	itemArray[115].wReserved = 0;

	/* 单井压力*/
	CString strPH1DP = "PH1_VALUE_PH1_DP";
	BSTR bstrPH1DP = strPH1DP.AllocSysString();
	itemArray[116].szAccessPath = L"";
	itemArray[116].szItemID = bstrPH1DP;
	itemArray[116].bActive = true; // active state
	itemArray[116].hClient = (OPCHANDLE)116; // our handle to item
	itemArray[116].dwBlobSize = 0; // no blob support
	itemArray[116].pBlob = NULL;
	itemArray[116].vtRequestedDataType = VT_R4;
	itemArray[116].wReserved = 0;

	/*瞬时流量*/
	CString strPH1F = "PH1_VALUE_PH1_F";
	BSTR bstrPH1F = strPH1F.AllocSysString();
	itemArray[117].szAccessPath = L"";
	itemArray[117].szItemID = bstrPH1F;
	itemArray[117].bActive = true; // active state
	itemArray[117].hClient = (OPCHANDLE)117; // our handle to item
	itemArray[117].dwBlobSize = 0; // no blob support
	itemArray[117].pBlob = NULL;
	itemArray[117].vtRequestedDataType = VT_R4;
	itemArray[117].wReserved = 0;

	/*累计流量*/
	CString strPH1S = "PH1_VALUE_PH1_S";
	BSTR bstrPH1S = strPH1S.AllocSysString();
	itemArray[118].szAccessPath = L"";
	itemArray[118].szItemID = bstrPH1S;
	itemArray[118].bActive = true; // active state
	itemArray[118].hClient = (OPCHANDLE)118; // our handle to item
	itemArray[118].dwBlobSize = 23; // no blob support
	itemArray[118].pBlob = NULL;
	itemArray[118].vtRequestedDataType = VT_R4;
	itemArray[118].wReserved = 0;

	/*第六个内存 第二个服务器句柄，指向数据读取*/
	hOPCServer2 = (OPCHANDLE*)CoTaskMemAlloc(119 * sizeof(OPCHANDLE));

	/*数据项添加*/
	hr = pIOPCItemMgt->AddItems(119, itemArray,
		(OPCITEMRESULT**)&pItemResult, (HRESULT**)&pErrors);
	ASSERT(pItemResult);
	ASSERT(pErrors);

	if (FAILED(hr))
	{
		cout << "未能为潘河组添加数据项..." << endl;
		if (pItemResult) CoTaskMemFree(pItemResult);
		if (pErrors) CoTaskMemFree(pErrors);
		CoTaskMemFree(&hOPCServer1);
		CoTaskMemFree(&hOPCServer2); //第六个内存释放
		CoTaskMemFree(&itemArray); //第五个内存释放
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
		pIOPCItemMgt = NULL;
		if (pIServer) pIServer->Release(); //第四个指针释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个指针释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {
		cout << "已经为潘河组添加数据项..." << endl;
	}

	/*为每个数据项添加服务器句柄*/

	hOPCServer2[0] = (OPCHANDLE)pItemResult[0].hServer;
	hOPCServer2[1] = (OPCHANDLE)pItemResult[1].hServer;
	hOPCServer2[2] = (OPCHANDLE)pItemResult[2].hServer;
	hOPCServer2[3] = (OPCHANDLE)pItemResult[3].hServer;
	hOPCServer2[4] = (OPCHANDLE)pItemResult[4].hServer;
	hOPCServer2[5] = (OPCHANDLE)pItemResult[5].hServer;
	hOPCServer2[6] = (OPCHANDLE)pItemResult[6].hServer;
	hOPCServer2[7] = (OPCHANDLE)pItemResult[7].hServer;
	hOPCServer2[8] = (OPCHANDLE)pItemResult[8].hServer;
	hOPCServer2[9] = (OPCHANDLE)pItemResult[9].hServer;
	hOPCServer2[10] = (OPCHANDLE)pItemResult[10].hServer;
	hOPCServer2[11] = (OPCHANDLE)pItemResult[11].hServer;
	hOPCServer2[12] = (OPCHANDLE)pItemResult[12].hServer;
	hOPCServer2[13] = (OPCHANDLE)pItemResult[13].hServer;
	hOPCServer2[14] = (OPCHANDLE)pItemResult[14].hServer;
	hOPCServer2[15] = (OPCHANDLE)pItemResult[15].hServer;
	hOPCServer2[16] = (OPCHANDLE)pItemResult[16].hServer;
	hOPCServer2[17] = (OPCHANDLE)pItemResult[17].hServer;
	hOPCServer2[18] = (OPCHANDLE)pItemResult[18].hServer;
	hOPCServer2[19] = (OPCHANDLE)pItemResult[19].hServer;
	hOPCServer2[20] = (OPCHANDLE)pItemResult[20].hServer;
	hOPCServer2[21] = (OPCHANDLE)pItemResult[21].hServer;
	hOPCServer2[22] = (OPCHANDLE)pItemResult[22].hServer;
	hOPCServer2[23] = (OPCHANDLE)pItemResult[23].hServer;
	hOPCServer2[24] = (OPCHANDLE)pItemResult[24].hServer;
	hOPCServer2[25] = (OPCHANDLE)pItemResult[25].hServer;
	hOPCServer2[26] = (OPCHANDLE)pItemResult[26].hServer;
	hOPCServer2[27] = (OPCHANDLE)pItemResult[27].hServer;
	hOPCServer2[28] = (OPCHANDLE)pItemResult[28].hServer;
	hOPCServer2[29] = (OPCHANDLE)pItemResult[29].hServer;
	hOPCServer2[30] = (OPCHANDLE)pItemResult[30].hServer;
	hOPCServer2[31] = (OPCHANDLE)pItemResult[31].hServer;
	hOPCServer2[32] = (OPCHANDLE)pItemResult[32].hServer;
	hOPCServer2[33] = (OPCHANDLE)pItemResult[33].hServer;
	hOPCServer2[34] = (OPCHANDLE)pItemResult[34].hServer;
	hOPCServer2[35] = (OPCHANDLE)pItemResult[35].hServer;
	hOPCServer2[36] = (OPCHANDLE)pItemResult[36].hServer;
	hOPCServer2[37] = (OPCHANDLE)pItemResult[37].hServer;
	hOPCServer2[38] = (OPCHANDLE)pItemResult[38].hServer;
	hOPCServer2[39] = (OPCHANDLE)pItemResult[39].hServer;
	hOPCServer2[40] = (OPCHANDLE)pItemResult[40].hServer;
	hOPCServer2[41] = (OPCHANDLE)pItemResult[41].hServer;
	hOPCServer2[42] = (OPCHANDLE)pItemResult[42].hServer;
	hOPCServer2[43] = (OPCHANDLE)pItemResult[43].hServer;
	hOPCServer2[44] = (OPCHANDLE)pItemResult[44].hServer;
	hOPCServer2[45] = (OPCHANDLE)pItemResult[45].hServer;
	hOPCServer2[46] = (OPCHANDLE)pItemResult[46].hServer;
	hOPCServer2[47] = (OPCHANDLE)pItemResult[47].hServer;
	hOPCServer2[48] = (OPCHANDLE)pItemResult[48].hServer;
	hOPCServer2[49] = (OPCHANDLE)pItemResult[49].hServer;
	hOPCServer2[50] = (OPCHANDLE)pItemResult[50].hServer;
	hOPCServer2[51] = (OPCHANDLE)pItemResult[51].hServer;
	hOPCServer2[52] = (OPCHANDLE)pItemResult[52].hServer;
	hOPCServer2[53] = (OPCHANDLE)pItemResult[53].hServer;
	hOPCServer2[54] = (OPCHANDLE)pItemResult[54].hServer;
	hOPCServer2[55] = (OPCHANDLE)pItemResult[55].hServer;
	hOPCServer2[56] = (OPCHANDLE)pItemResult[56].hServer;
	hOPCServer2[57] = (OPCHANDLE)pItemResult[57].hServer;
	hOPCServer2[58] = (OPCHANDLE)pItemResult[58].hServer;
	hOPCServer2[59] = (OPCHANDLE)pItemResult[59].hServer;
	hOPCServer2[60] = (OPCHANDLE)pItemResult[60].hServer;
	hOPCServer2[61] = (OPCHANDLE)pItemResult[61].hServer;
	hOPCServer2[62] = (OPCHANDLE)pItemResult[62].hServer;
	hOPCServer2[63] = (OPCHANDLE)pItemResult[63].hServer;
	hOPCServer2[64] = (OPCHANDLE)pItemResult[64].hServer;
	hOPCServer2[65] = (OPCHANDLE)pItemResult[65].hServer;
	hOPCServer2[66] = (OPCHANDLE)pItemResult[66].hServer;
	hOPCServer2[67] = (OPCHANDLE)pItemResult[67].hServer;
	hOPCServer2[68] = (OPCHANDLE)pItemResult[68].hServer;
	hOPCServer2[69] = (OPCHANDLE)pItemResult[69].hServer;
	hOPCServer2[70] = (OPCHANDLE)pItemResult[70].hServer;
	hOPCServer2[71] = (OPCHANDLE)pItemResult[71].hServer;
	hOPCServer2[72] = (OPCHANDLE)pItemResult[72].hServer;
	hOPCServer2[73] = (OPCHANDLE)pItemResult[73].hServer;
	hOPCServer2[74] = (OPCHANDLE)pItemResult[74].hServer;
	hOPCServer2[75] = (OPCHANDLE)pItemResult[75].hServer;
	hOPCServer2[76] = (OPCHANDLE)pItemResult[76].hServer;
	hOPCServer2[77] = (OPCHANDLE)pItemResult[77].hServer;
	hOPCServer2[78] = (OPCHANDLE)pItemResult[78].hServer;
	hOPCServer2[79] = (OPCHANDLE)pItemResult[79].hServer;
	hOPCServer2[80] = (OPCHANDLE)pItemResult[80].hServer;
	hOPCServer2[81] = (OPCHANDLE)pItemResult[81].hServer;
	hOPCServer2[82] = (OPCHANDLE)pItemResult[82].hServer;
	hOPCServer2[83] = (OPCHANDLE)pItemResult[83].hServer;
	hOPCServer2[84] = (OPCHANDLE)pItemResult[84].hServer;
	hOPCServer2[85] = (OPCHANDLE)pItemResult[85].hServer;
	hOPCServer2[86] = (OPCHANDLE)pItemResult[86].hServer;
	hOPCServer2[87] = (OPCHANDLE)pItemResult[87].hServer;
	hOPCServer2[88] = (OPCHANDLE)pItemResult[88].hServer;
	hOPCServer2[89] = (OPCHANDLE)pItemResult[89].hServer;
	hOPCServer2[90] = (OPCHANDLE)pItemResult[90].hServer;
	hOPCServer2[91] = (OPCHANDLE)pItemResult[91].hServer;
	hOPCServer2[92] = (OPCHANDLE)pItemResult[92].hServer;
	hOPCServer2[93] = (OPCHANDLE)pItemResult[93].hServer;
	hOPCServer2[94] = (OPCHANDLE)pItemResult[94].hServer;
	hOPCServer2[95] = (OPCHANDLE)pItemResult[95].hServer;
	hOPCServer2[96] = (OPCHANDLE)pItemResult[96].hServer;
	hOPCServer2[97] = (OPCHANDLE)pItemResult[97].hServer;
	hOPCServer2[98] = (OPCHANDLE)pItemResult[98].hServer;
	hOPCServer2[99] = (OPCHANDLE)pItemResult[99].hServer;
	hOPCServer2[100] = (OPCHANDLE)pItemResult[100].hServer;
	hOPCServer2[101] = (OPCHANDLE)pItemResult[101].hServer;
	hOPCServer2[102] = (OPCHANDLE)pItemResult[102].hServer;
	hOPCServer2[103] = (OPCHANDLE)pItemResult[103].hServer;
	hOPCServer2[104] = (OPCHANDLE)pItemResult[104].hServer;
	hOPCServer2[105] = (OPCHANDLE)pItemResult[105].hServer;
	hOPCServer2[106] = (OPCHANDLE)pItemResult[106].hServer;
	hOPCServer2[107] = (OPCHANDLE)pItemResult[107].hServer;
	hOPCServer2[108] = (OPCHANDLE)pItemResult[108].hServer;
	hOPCServer2[109] = (OPCHANDLE)pItemResult[109].hServer;
	hOPCServer2[110] = (OPCHANDLE)pItemResult[110].hServer;
	hOPCServer2[111] = (OPCHANDLE)pItemResult[111].hServer;
	hOPCServer2[112] = (OPCHANDLE)pItemResult[112].hServer;
	hOPCServer2[113] = (OPCHANDLE)pItemResult[113].hServer;
	hOPCServer2[114] = (OPCHANDLE)pItemResult[114].hServer;
	hOPCServer2[115] = (OPCHANDLE)pItemResult[115].hServer;
	hOPCServer2[116] = (OPCHANDLE)pItemResult[116].hServer;
	hOPCServer2[117] = (OPCHANDLE)pItemResult[117].hServer;
	hOPCServer2[118] = (OPCHANDLE)pItemResult[118].hServer;

	/*同步读取数据*/
	//IOPCSyncIO *pIOPCSyncIO = NULL; //第十个指针，指向同步读取
	//hr = pIOPCItemMgt->QueryInterface(IID_IOPCSyncIO, /*OUT*/ (void**)&pIOPCSyncIO);
	//if (FAILED(hr)) {
	//	cout << "未能获取到接口IOPCSyncIO" << endl;
	//	if (pItemResult) CoTaskMemFree(pItemResult);
	//	if (pErrors) CoTaskMemFree(pErrors);
	//	if (pIOPCSyncIO) CoTaskMemFree(pIOPCSyncIO);
	//	CoTaskMemFree(&hOPCServer1);
	//	CoTaskMemFree(&hOPCServer2);
	//	if (pIServer) pIServer->Release();
	//	pIServer = NULL;
	//	if (pIUnknown) pIUnknown->Release();
	//	pIUnknown = NULL;
	//	if (pIEnumGUID) pIEnumGUID->Release();
	//	pIEnumGUID = NULL;
	//	if (pIServerList) pIServerList->Release();
	//	pIServerList = NULL;
	//	if (pIOPCItemMgt) pIOPCItemMgt->Release();
	//	pIOPCItemMgt = NULL;
	//	return 1;
	//}
	//else if (SUCCEEDED(hr)) {
	//	ASSERT(pIOPCSyncIO);
	//	cout << "已获取到接口IOPCSyncIO" << endl;
	//}

	//DWORD dwItemNumber = 1;
	//OPCITEMSTATE *pItemValue = NULL; //第十二个指针，指向读取OPC SERVER的值

	//hr = pIOPCSyncIO->Read(OPC_DS_CACHE, dwItemNumber, hOPCServer2, &pItemValue, &pErrors);

	//if (FAILED(hr))
	//{
	//	cout << "未能读取到OPC服务器端相对应的数据，请查看数据源是否有误" << endl;
	//	//  if (m_hServer) CoTaskMemFree (m_hServer);
	//	if (pItemResult) CoTaskMemFree(pItemResult);
	//	if (pErrors) CoTaskMemFree(pErrors);
	//	if (pIOPCSyncIO) CoTaskMemFree(pIOPCSyncIO);
	//	CoTaskMemFree(&hOPCServer1);
	//	CoTaskMemFree(&hOPCServer2);
	//	CoTaskMemFree(&pItemValue);
	//	if (pIServer) pIServer->Release();
	//	pIServer = NULL;
	//	if (pIUnknown) pIUnknown->Release();
	//	pIUnknown = NULL;
	//	if (pIEnumGUID) pIEnumGUID->Release();
	//	pIEnumGUID = NULL;
	//	if (pIServerList) pIServerList->Release();
	//	pIServerList = NULL;
	//	if (pIOPCItemMgt) pIOPCItemMgt->Release();
	//	pIOPCItemMgt = NULL;
	//	VariantClear(&pItemValue[0].vDataValue);
	//	return 1;   //同步读数据时出错
	//}
	//else if (SUCCEEDED(hr)){
	//  ASSERT(hr);
	//	cout << "已读取到OPC服务器端的数据" << endl;
	//	cout << V_R4(&pItemValue[0].vDataValue) << endl;
	//}

	/*组更新状态*/
	hr = pIOPCItemMgt->QueryInterface(IID_IOPCGroupStateMgt, /*OUT*/(void**)&pIOPCGroupStateMgt); //得到第十个指针
	ASSERT(pIOPCGroupStateMgt);

	if (FAILED(hr))
	{
		cout << "获取IOPCGroupStateMgt接口失败..." << endl;
		if (pItemResult) CoTaskMemFree(pItemResult);
		if (pErrors) CoTaskMemFree(pErrors);
		CoTaskMemFree(&hOPCServer1);
		CoTaskMemFree(&hOPCServer2); //第六个内存释放
		CoTaskMemFree(&itemArray); //第五个内存释放
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIOPCGroupStateMgt) pIOPCGroupStateMgt->Release(); //第十个指针释放
		pIOPCGroupStateMgt = NULL;
		if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
		pIOPCItemMgt = NULL;
		if (pIServer) pIServer->Release(); //第四个指针释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个指针释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {
		cout << "已获取到IOPCGroupStateMgt接口..." << endl;
	}

	DWORD  dwRevUpdateRate = 10;
	BOOL bActivateGroup = TRUE;

	hr = pIOPCGroupStateMgt->SetState(/*[in] RequestedUpdateRate*/ NULL, /*[out] RevisedUpdateRate */ &dwRevUpdateRate, /*[in] ActiveFlag for Group */ &bActivateGroup, /*[in] TimeBias*/ NULL, /*[in] PercentDeadband*/ NULL, /*[in] LCID*/ NULL, NULL);	

	/*设置回调*/
	CComObject<COPCDataCallback> *pCOPCDataCallback;
	CComObject<COPCDataCallback>::CreateInstance(&pCOPCDataCallback);

	LPUNKNOWN pCbUnk;
	pCbUnk = pCOPCDataCallback->GetUnknown();
	DWORD dwCookie;

	HRESULT hRes = AtlAdvise(pIOPCGroupStateMgt, pCbUnk, IID_IOPCDataCallback, &dwCookie);

	if (FAILED(hRes)) {
		cout << "数据订阅回调设置失败..." << endl;
		if (pItemResult) CoTaskMemFree(pItemResult);
		if (pErrors) CoTaskMemFree(pErrors);
		CoTaskMemFree(&hOPCServer1);
		CoTaskMemFree(&hOPCServer2); //第六个内存释放
		CoTaskMemFree(&itemArray); //第五个内存释放
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIOPCGroupStateMgt) pIOPCGroupStateMgt->Release(); //第十个指针释放
		pIOPCGroupStateMgt = NULL;
		if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
		pIOPCItemMgt = NULL;
		if (pIServer) pIServer->Release(); //第四个指针释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个指针释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hRes)) {
		cout << "数据订阅回调设置成功..." << endl;
	}

	/*异步读取数据*/
	
	hr = pIOPCItemMgt->QueryInterface(IID_IOPCAsyncIO2, /*OUT*/(void**)&pIOPCAsyncIO2); //得到第十一个指针
	ASSERT(pIOPCAsyncIO2);

	if (FAILED(hr)) {
		cout << "未能获取到接口IOPCSyncIO2..." << endl;
		if (pItemResult) CoTaskMemFree(pItemResult);
		if (pErrors) CoTaskMemFree(pErrors);
		CoTaskMemFree(&hOPCServer1);
		CoTaskMemFree(&hOPCServer2); //第六个内存释放
		CoTaskMemFree(&itemArray); //第五个内存释放
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIOPCAsyncIO2) pIOPCAsyncIO2->Release(); //第十一个指针释放
		pIOPCAsyncIO2 = NULL;
		if (pIOPCGroupStateMgt) pIOPCGroupStateMgt->Release(); //第十个指针释放
		pIOPCGroupStateMgt = NULL;
		if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
		pIOPCItemMgt = NULL;
		if (pIServer) pIServer->Release(); //第四个指针释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个指针释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;
	}
	else if (SUCCEEDED(hr)) {		
		cout << "已获取到接口IOPCSyncIO2..." << endl;
	}

	DWORD dwItemNumber = 119;
	DWORD dwCancelID;
	OPCITEMSTATE *pItemValue = NULL; //第十二个指针，指向读取OPC SERVER的值

	hr = pIOPCAsyncIO2->Read(dwItemNumber, hOPCServer2, 1, &dwCancelID, &pErrors); //异步读取，数据在COPCDataCallback中返回

	if (FAILED(hr)) {
		cout << "未能读取到OPC服务器端相对应的数据，请查看数据源是否有误..." << endl;
		VariantClear(&pItemValue[0].vDataValue);
		if (pItemResult) CoTaskMemFree(pItemResult);
		if (pErrors) CoTaskMemFree(pErrors);
		CoTaskMemFree(&hOPCServer1);
		CoTaskMemFree(&hOPCServer2); //第六个内存释放
		CoTaskMemFree(&itemArray); //第五个内存释放
		CoTaskMemFree(&clsid); //第四个内存释放
		CoTaskMemFree(&catID); //第三个内存释放
		CoTaskMemFree(&mqi); //第二个内存释放
		CoTaskMemFree(&si); //第一个内存释放
		if (pIOPCAsyncIO2) pIOPCAsyncIO2->Release(); //第十一个指针释放
		pIOPCAsyncIO2 = NULL;
		if (pIOPCGroupStateMgt) pIOPCGroupStateMgt->Release(); //第十个指针释放
		pIOPCGroupStateMgt = NULL;
		if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
		pIOPCItemMgt = NULL;
		if (pIServer) pIServer->Release(); //第四个指针释放
		pIServer = NULL;
		if (pIUnknown) pIUnknown->Release(); //第三个指针释放
		pIUnknown = NULL;
		if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
		pIEnumGUID = NULL;
		if (pIServerList) pIServerList->Release(); //第一个指针释放
		pIServerList = NULL;
		return 1;   //读数据时出错
	}
	else if (SUCCEEDED(hr)) {
		cout << "已获取到数据..." << endl;
	}	

	system("pause>nul");

	if (pItemResult) CoTaskMemFree(pItemResult);
	if (pErrors) CoTaskMemFree(pErrors);
	CoTaskMemFree(&hOPCServer1);
	CoTaskMemFree(&hOPCServer2); //第六个内存释放
	CoTaskMemFree(&itemArray); //第五个内存释放
	CoTaskMemFree(&clsid); //第四个内存释放
	CoTaskMemFree(&catID); //第三个内存释放
	CoTaskMemFree(&mqi); //第二个内存释放
	CoTaskMemFree(&si); //第一个内存释放
	if (pIOPCAsyncIO2) pIOPCAsyncIO2->Release(); //第十一个指针释放
	pIOPCAsyncIO2 = NULL;
	if (pIOPCGroupStateMgt) pIOPCGroupStateMgt->Release(); //第十个指针释放
	pIOPCGroupStateMgt = NULL;
	if (pIOPCItemMgt) pIOPCItemMgt->Release(); //第五个指针释放
	pIOPCItemMgt = NULL;
	if (pIServer) pIServer->Release(); //第四个指针释放
	pIServer = NULL;
	if (pIUnknown) pIUnknown->Release(); //第三个指针释放
	pIUnknown = NULL;
	if (pIEnumGUID) pIEnumGUID->Release(); //第二个指针释放
	pIEnumGUID = NULL;
	if (pIServerList) pIServerList->Release(); //第一个指针释放
	pIServerList = NULL;


	//pIOPCSyncIO->Release();
	//pCOPCDataCallback->Release();

	CoUninitialize();	
	return 0;
}