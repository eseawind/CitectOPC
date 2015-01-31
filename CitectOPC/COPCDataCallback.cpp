#include "stdafx.h"
#include <afxdisp.h>
#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <set>
#include "COPCDataCallback.h"

int changeFlag = 1,readFlag,writeFlag;
std::vector<int> changeNum;

float value[119];
float changevalues[119] = {0};
std::string groupName[16] = { "PH63_10", "PH73_11", "PH74_08", "PH64_01", "PH86_01", "PH65_11", "PH55_09", "PH76_03", "PH66_01", "PH77_08",
"PE_079", "PE_112", "PH73-11", "PH45-05", "PH35-07", "PH1-02" };

CString itemName[119] = { "PH1_VALUE_PH63_HT", "PH1_VALUE_PH63_HP", "PH1_VALUE_PH63_HF", "PH1_VALUE_PH63_HS",
			              "PH1_VALUE_PH63_DT", "PH1_VALUE_PH63_DP", "PH1_VALUE_PH63_DF", "PH1_VALUE_PH63_DS",
						  "PH1_VALUE_PH7311_HT", "PH1_VALUE_PH7311_HP", "PH1_VALUE_PH7311_HF", "PH1_VALUE_PH7311_HS",
						  "PH1_VALUE_PH7311_DT", "PH1_VALUE_PH7311_DP", "PH1_VALUE_PH7311_DF", "PH1_VALUE_PH7311_DS", 
						  "PH1_VALUE_PH74_HT", "PH1_VALUE_PH74_HP",	"PH1_VALUE_PH74_HF", "PH1_VALUE_PH74_HS", 
						  "PH1_VALUE_PH74_DT", "PH1_VALUE_PH74_DP", "PH1_VALUE_PH74_DF", "PH1_VALUE_PH74_DS", 
						  "PH1_VALUE_PH64_HT", "PH1_VALUE_PH64_HP", "PH1_VALUE_PH64_HF", "PH1_VALUE_PH64_HS", 
						  "PH1_VALUE_PH64_DT", "PH1_VALUE_PH64_DP", "PH1_VALUE_PH64_DF", "PH1_VALUE_PH64_DS", 
						  "PH1_VALUE_PH86_HT", "PH1_VALUE_PH86_HP", "PH1_VALUE_PH86_HF", "PH1_VALUE_PH86_HS", 
						  "PH1_VALUE_PH86_DT", "PH1_VALUE_PH86_DP",	"PH1_VALUE_PH86_DF", "PH1_VALUE_PH86_DS", 
						  "PH1_VALUE_PH65_HT", "PH1_VALUE_PH65_HP", "PH1_VALUE_PH65_HF", "PH1_VALUE_PH65_HS", 
						  "PH1_VALUE_PH65_DT", "PH1_VALUE_PH65_DP", "PH1_VALUE_PH65_DF", "PH1_VALUE_PH65_DS",
						  "PH1_VALUE_PH55_HT", "PH1_VALUE_PH55_HP", "PH1_VALUE_PH55_HF", "PH1_VALUE_PH55_HS", 
						  "PH1_VALUE_PH55_DT", "PH1_VALUE_PH55_DP", "PH1_VALUE_PH55_DF", "PH1_VALUE_PH55_DS", 
						  "PH1_VALUE_PH76_HT", "PH1_VALUE_PH76_HP", "PH1_VALUE_PH76_HF", "PH1_VALUE_PH76_HS", 
						  "PH1_VALUE_PH76_DT", "PH1_VALUE_PH76_DP", "PH1_VALUE_PH76_DF", "PH1_VALUE_PH76_DS", 
						  "PH1_VALUE_PH66_HT", "PH1_VALUE_PH66_HP", "PH1_VALUE_PH66_HF", "PH1_VALUE_PH66_HS", 
						  "PH1_VALUE_PH66_DT", "PH1_VALUE_PH66_DP", "PH1_VALUE_PH66_DF", "PH1_VALUE_PH66_DS", 
						  "PH1_VALUE_PH77_HT", "PH1_VALUE_PH77_HP", "PH1_VALUE_PH77_HF", "PH1_VALUE_PH77_HS", 
						  "PH1_VALUE_PH77_DT", "PH1_VALUE_PH77_DP", "PH1_VALUE_PH77_DF", "PH1_VALUE_PH77_DS", 
						  "PH1_VALUE_PE79_HT", "PH1_VALUE_PE79_HP", "PH1_VALUE_PE79_HF", "PH1_VALUE_PE79_HS", 
						  "PH1_VALUE_PE79_DT", "PH1_VALUE_PE79_DP", "PH1_VALUE_PE79_DF", "PH1_VALUE_PE79_DS",
						  "PH1_VALUE_PE112_HT", "PH1_VALUE_PE112_HP", "PH1_VALUE_PE112_HF", "PH1_VALUE_PE112_HS", 
						  "PH1_VALUE_PE112_DT",	"PH1_VALUE_PE112_DP", "PH1_VALUE_PE112_DF", "PH1_VALUE_PE112_DS", 
						  "PH1_VALUE_PH73_HT", "PH1_VALUE_PH73_HP", "PH1_VALUE_PH73_HF", "PH1_VALUE_PH73_HS", 
						  "PH1_VALUE_PH73_DT", "PH1_VALUE_PH73_DP", "PH1_VALUE_PH73_DF", "PH1_VALUE_PH73_DS", 
						  "PH1_VALUE_PH45_HP", "PH1_VALUE_PH45_DT", "PH1_VALUE_PH45_DP", "PH1_VALUE_PH45_F", "PH1_VALUE_PH45_S", 
						  "PH1_VALUE_PH35_HP", "PH1_VALUE_PH35_DT", "PH1_VALUE_PH35_DP", "PH1_VALUE_PH35_F", "PH1_VALUE_PH35_S",
						  "PH1_VALUE_PH1_HP", "PH1_VALUE_PH1_DT", "PH1_VALUE_PH1_DP", "PH1_VALUE_PH1_F", "PH1_VALUE_PH1_S" };

CString quility[119];
CString timestamp[119];
float readValue[119];
CString readQulity[119];
CString readTS[119];
HRESULT writeRes[119];

CComModule _Module;

CString GetQualityText(UINT qnr)
{
	CString qstr;

	switch (qnr)
	{
	case OPC_QUALITY_BAD:           qstr = "BAD";
		break;
	case OPC_QUALITY_UNCERTAIN:     qstr = "UNCERTAIN";
		break;
	case OPC_QUALITY_GOOD:          qstr = "GOOD";
		break;
	case OPC_QUALITY_NOT_CONNECTED: qstr = "NOT_CONNECTED";
		break;
	case OPC_QUALITY_DEVICE_FAILURE:qstr = "DEVICE_FAILURE";
		break;
	case OPC_QUALITY_SENSOR_FAILURE:qstr = "SENSOR_FAILURE";
		break;
	case OPC_QUALITY_LAST_KNOWN:    qstr = "LAST_KNOWN";
		break;
	case OPC_QUALITY_COMM_FAILURE:  qstr = "COMM_FAILURE";
		break;
	case OPC_QUALITY_OUT_OF_SERVICE:qstr = "OUT_OF_SERVICE";
		break;
	case OPC_QUALITY_LAST_USABLE:   qstr = "LAST_USABLE";
		break;
	case OPC_QUALITY_SENSOR_CAL:    qstr = "SENSOR_CAL";
		break;
	case OPC_QUALITY_EGU_EXCEEDED:  qstr = "EGU_EXCEEDED";
		break;
	case OPC_QUALITY_SUB_NORMAL:    qstr = "SUB_NORMAL";
		break;
	case OPC_QUALITY_LOCAL_OVERRIDE:qstr = "LOCAL_OVERRIDE";
		break;
	default:                        qstr = "UNKNOWN ERROR";
	}
	return qstr;
};

void Data2Text(int hItem, float *values) {
	SYSTEMTIME st = { 0 };
	GetLocalTime(&st);

	if (hItem < 105){
		std::ofstream outFile;
		outFile.open("valueblock.txt", std::ios::app);

		outFile << groupName[hItem/8 - 1] << ",";
		outFile << st.wYear << "-" << st.wMonth << "-" << st.wDay << " ";
		outFile << st.wHour << ":" << st.wMinute << ":" << st.wSecond << ",";

		for (int i = hItem-8; i < hItem; i++) {
			if (i != hItem-1) {
				outFile << values[i] << ",";
			}
			else {
				outFile << values[i];
			}

		}

		outFile << "\n";
		outFile.close();
		std::cout << groupName[hItem/8 - 1] << "数据已写入完毕" << std::endl;
	}
	else
	{
		std::ofstream outFile;
		outFile.open("individualwell.txt", std::ios::app);
		
		outFile << groupName[(hItem - 104) / 5 + 12] << ",";
		outFile << st.wYear << "-" << st.wMonth << "-" << st.wDay << " ";
		outFile << st.wHour << ":" << st.wMinute << ":" << st.wSecond << ",";

		for (int i = hItem - 5; i < hItem; i++) {
			if (i !=  hItem-1) {
				outFile << values[i] << ",";
			}
			else {
				outFile << values[i];
			}

		}

		outFile << "\n";
		outFile.close();
		std::cout << groupName[(hItem-104)/5 + 12] << "数据已写入完毕" << std::endl;
	}
};

STDMETHODIMP COPCDataCallback::OnDataChange(  // OnDataChange notifications
	DWORD dwTransID,   // 0 for normal OnDataChange events, non-zero for Refreshes
	OPCHANDLE hGroup,   // client group handle
	HRESULT hrMasterQuality, // S_OK if all qualities are GOOD, otherwise S_FALSE
	HRESULT hrMasterError,  // S_OK if all errors are S_OK, otherwise S_FALSE
	DWORD dwCount,    // number of items in the lists that follow
	OPCHANDLE *phClientItems, // item client handles
	VARIANT *pvValues,   // item data
	WORD *pwQualities,   // item qualities
	FILETIME *pftTimeStamps, // item timestamps
	HRESULT *pErrors)   // item errors 
{
	std::cout << std::endl;
	std::cout << "数据第" << changeFlag << "次变更：" << std::endl;
	DWORD i;
	for (i = 0; i<dwCount; i++)
	{
		value[i] = pvValues[i].fltVal;
		changevalues[(int)phClientItems[i]] = value[i];
		quility[i] = GetQualityText(pwQualities[i]);
		timestamp[i] = COleDateTime(pftTimeStamps[i]).Format();
		std::wcout << itemName[(int)phClientItems[i]].GetString() << "的值为：" << value[i] <<std::endl;

		changeNum.push_back((int)phClientItems[i]);
	}

	std::vector<int> itemNum;

	std::vector<int>::iterator it;
	for (it = changeNum.begin(); it != changeNum.end(); it++) {
		int f = 0;
		if (*it >= 0 && *it < 8) {
			f = 8;
		}
		else if (*it >= 8 && *it < 16) {
			f = 16;
		}
		else if (*it >= 16 && *it < 24) {
			f = 24;
		}
		else if (*it >= 24 && *it < 32) {
			f = 32;
		}
		else if (*it >= 32 && *it < 40) {
			f = 40;
		}
		else if (*it >= 40 && *it < 48) {
			f = 48;
		}
		else if (*it >= 48 && *it < 56) {
			f = 56;
		}
		else if (*it >= 56 && *it < 64) {
			f = 64;
		}
		else if (*it >= 64 && *it < 72) {
			f = 72;
		}
		else if (*it >= 72 && *it < 80) {
			f = 80;
		}
		else if (*it >= 80 && *it < 88) {
			f = 88;
		}
		else if (*it >= 88 && *it < 96) {
			f = 96;
		}
		else if (*it >= 96 && *it < 104) {
			f = 104;
		}
		else if (*it >= 104 && *it < 109) {
			f = 109;
		}
		else if (*it >= 109 && *it < 114) {
			f = 114;
		}
		else if (*it >= 114 && *it < 119) {
			f = 119;
		}

		itemNum.push_back(f);
	}
	

	std::set<int> si(itemNum.begin(), itemNum.end());
	std::vector<int> changeNew(si.begin(), si.end());

	for (std::vector<int>::iterator it = changeNew.begin(); it != changeNew.end(); it++) {
		Data2Text(*it, changevalues);
	}

	std::cout << "数据变更完毕..." << std::endl;
	changeFlag++;

	std::vector <int>().swap(changeNew);
	std::set <int>().swap(si);
	std::vector <int>().swap(changeNum);
	return S_OK;
};

STDMETHODIMP COPCDataCallback::OnReadComplete( // OnReadComplete notifications
	DWORD dwTransID,   // Transaction ID returned by the server when the read was initiated
	OPCHANDLE hGroup,   // client group handle
	HRESULT hrMasterQuality, // S_OK if all qualities are GOOD, otherwise S_FALSE
	HRESULT hrMasterError,  // S_OK if all errors are S_OK, otherwise S_FALSE
	DWORD dwCount,    // number of items in the lists that follow
	OPCHANDLE *phClientItems, // item client handles
	VARIANT *pvValues,   // item data
	WORD *pwQualities,   // item qualities
	FILETIME *pftTimeStamps, // item timestamps
	HRESULT *pErrors)   // item errors 
{	
	std::cout << std::endl;
	std::cout << "数据获取中..." << std::endl;
	readFlag = 1;
	if (pErrors[0] == S_OK)
	{
		DWORD i;
		for (i = 0; i<dwCount; i++)
		{
			readValue[i] = pvValues[i].fltVal;
			changevalues[phClientItems[i]] = readValue[i];
			readQulity[i] = GetQualityText(pwQualities[i]);
			readTS[i] = COleDateTime(pftTimeStamps[i]).Format();
			if (phClientItems[i] == (OPCHANDLE)0) {
				std::cout << " PH63_10阀组（2#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)8) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH73_11阀组（3#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)16) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH74_08阀组（4#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)24) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH64_01阀组（5#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)32) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH86_01阀组（7#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)40) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH65_11阀组（8#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)48) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH55_09阀组（9#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)56) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH76_03阀组（10#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)64) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH66_01阀组（11#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)72) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH77_08阀组（12#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)80) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PE_079阀组（14#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)88) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PE_112阀组（15#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)96) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH73-11阀组（3#）" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)104) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH45-05单井轮换" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)109) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH35-07单井轮换" << std::endl;
			}
			else if (phClientItems[i] == (OPCHANDLE)114) {
				Data2Text((int)phClientItems[i], readValue);
				std::cout << " PH1-02单井轮换" << std::endl;
			}
			
			else {
				
			}
			
			std::wcout << itemName[(int)phClientItems[i]].GetString() << "的值为：" << readValue[i] << std::endl;
		}
		Data2Text(119, readValue);
	}
	else
	{
		CString readQuality = GetQualityText(pErrors[0]);
	}
	std::cout << "数据已全部读完..." << std::endl;

	return S_OK;
};

STDMETHODIMP COPCDataCallback::OnWriteComplete( // OnWriteComplete notifications
	DWORD dwTransID,   // Transaction ID returned by the server when the write was initiated
	OPCHANDLE hGroup,   // client group handle
	HRESULT hrMasterError,  // S_OK if all errors are S_OK, otherwise S_FALSE
	DWORD dwCount,    // number of items in the lists that follow
	OPCHANDLE *phClientItems, // item client handles
	HRESULT *pErrors)   // item errors 
{
	writeFlag = 1;
	writeRes[0] = pErrors[0];
	return S_OK;
};