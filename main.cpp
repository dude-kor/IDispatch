#include <ole2.h>

int main() {
	HRESULT CoInitialize(void/*IN LPVOID pvReserved*/);
	
	/*Example*/
	CLSID clsid;	// �ҷ��� ��ü�� CLSID
	IDispatch* pDispatch;	// �ҷ��� ��ü�� ������ ��
	CLSIDFromString(OLESTR("{00000000-0000-0000-0000-000000000000}"), &clsid);
	CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)&pDispatch);

	return 0;
}