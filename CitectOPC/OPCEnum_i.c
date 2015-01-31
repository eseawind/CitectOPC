/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Nov 20 22:30:48 2005
 */
/* Compiler settings for opcenum.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID LIBID_OpcEnumLib = {0x13486D43,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}};


//const IID IID_IOPCServerList2 = {0x9DD0B56C,0xAD9E,0x43EE,{0x83,0x05,0x48,0x7F,0x31,0x88,0xBF,0x7A}};


//const IID IID_IOPCEnumGUID = {0x55C382C8,0x21C7,0x4E88,{0x96,0xC1,0xBE,0xCF,0xB1,0xE3,0xF4,0x83}};


//const IID IID_IOPCServerList = {0x13486D50,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}};

#ifndef _CONST_IID_IEnumGUID
#define _CONST_IID_IEnumGUID
const IID IID_IEnumGUID = {0x0002E000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
#endif


const CLSID CLSID_OpcServerList = {0x13486D51,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}};


#ifdef __cplusplus
}
#endif

