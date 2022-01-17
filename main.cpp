#include <ole2.h>

int main() {
	HRESULT CoInitialize(void/*IN LPVOID pvReserved*/);
	
	/*Example*/
	CLSID clsid;	// 불러올 객체의 CLSID
	IDispatch* pDispatch;	// 불러온 객체가 보관될 곳
	CLSIDFromString(OLESTR("{00000000-0000-0000-0000-000000000000}"), &clsid);
	CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)&pDispatch);

	return 0;
}