
#include <windows.h>
#include <exdisp.h>
#include <exdispid.h>
#include <commctrl.h>
#pragma comment(lib,"Comctl32")
// Définir les messages pour notre application:
#define BEFORENAVIGATE2       WM_USER
#define DOWNLOADBEGIN         WM_USER+1
#define DOWNLOADCOMPLETE      WM_USER+2
#define NAVIGATECOMPLETE2     WM_USER+3
#define DOCUMENTCOMPLETE      WM_USER+4
#define COMMANDSTATECHANGE    WM_USER+5

#pragma warning(disable:26495)
#pragma warning(disable:6054)
#pragma warning(disable:6387)
#pragma warning(disable:6031)
#pragma warning(disable:8012)
#pragma warning(disable:28251)
#pragma warning(disable:9035)
#pragma warning(disable:4616)
#pragma warning(disable:4244)
#pragma warning(disable:4312)

class Evenem : public IDispatch
{
	private:
	long ref;
	HWND fenetre;
	BSTR url;

	public:
		Evenem(HWND fenet){fenetre=fenet;}
		~Evenem(){SysFreeString(url);}
		STDMETHODIMP QueryInterface(REFIID iid, void ** ppvObject){	if (iid==IID_IUnknown || iid==IID_IDispatch || iid==DIID_DWebBrowserEvents2){*ppvObject=this;AddRef(); return S_OK;} else return E_NOINTERFACE;	}
		ULONG STDMETHODCALLTYPE AddRef()	{ 	return InterlockedIncrement(&ref);  	}
		ULONG STDMETHODCALLTYPE Release(){	int tmp = InterlockedDecrement(&ref);	if (tmp==0) delete this; 	return tmp;	}
		HRESULT STDMETHODCALLTYPE GetTypeInfoCount(unsigned int FAR* pctinfo){return E_NOTIMPL;}
		HRESULT STDMETHODCALLTYPE GetTypeInfo(unsigned int iTInfo, LCID  lcid, ITypeInfo FAR* FAR*  ppTInfo){return E_NOTIMPL;}
		HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames, LCID lcid, DISPID FAR* rgDispId){ return E_NOTIMPL; }
		HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pDispParams, VARIANT FAR* parResult, EXCEPINFO FAR* pExcepInfo, unsigned int FAR* puArgErr)	{
			IUnknown *pIUnk;
			VARIANT  *vurl ;
			if (!pDispParams) return E_INVALIDARG;  
			switch (dispIdMember) {
				case DISPID_BEFORENAVIGATE2:{pIUnk = pDispParams->rgvarg[6].pdispVal;SendMessage(fenetre, BEFORENAVIGATE2, (WPARAM)pIUnk, 0);}break;
				case DISPID_DOWNLOADBEGIN: {SendMessage(fenetre,DOWNLOADBEGIN,0,0);}break;
				case DISPID_DOWNLOADCOMPLETE:{SendMessage(fenetre,DOWNLOADCOMPLETE,0,0);}break;
				case DISPID_NAVIGATECOMPLETE2: {
					pIUnk=pDispParams->rgvarg[1].pdispVal;
					vurl= pDispParams->rgvarg[0].pvarVal;
					url = vurl->bstrVal;
					SendMessage(fenetre,NAVIGATECOMPLETE2,(WPARAM)pIUnk,(LPARAM)url);
				}break;
				case DISPID_DOCUMENTCOMPLETE:{pIUnk=pDispParams->rgvarg[1].pdispVal;SendMessage(fenetre,DOCUMENTCOMPLETE,(WPARAM)pIUnk,0);}break;
				case DISPID_COMMANDSTATECHANGE: {
					long command =pDispParams->rgvarg[1].lVal;
					VARIANT_BOOL etat=pDispParams->rgvarg[0].boolVal;
					SendMessage(fenetre,COMMANDSTATECHANGE,(WPARAM)command,(LPARAM)etat);
				}break;
				case DISPID_NEWWINDOW2:{pDispParams->rgvarg[0].pvarVal->vt = VT_BOOL;pDispParams->rgvarg[0].pvarVal->boolVal = VARIANT_TRUE;}break;
				default:
					break;
				}	
		  return S_OK;
			}
};
IWebBrowser2   *pIWeb;
WNDPROC OldEditProc;
HWND hConteneur,hAdresse;
LRESULT CALLBACK ProcedureSousClassementEdit(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam){
	switch (message){
	case WM_CHAR: if (wParam == VK_RETURN) return 0;break;
	case WM_KEYDOWN:
		if (wParam == VK_RETURN){
			WCHAR url[256];
			char buff[256];
			GetWindowText(hAdresse,buff,256);
			MultiByteToWideChar (CP_ACP, 0,buff, -1, url, 256);
			pIWeb->Navigate(url,0,0,0,0);
			return 0;
		}
		break;
	default:
		break;
	}
	return CallWindowProc(OldEditProc, hwnd, message, wParam, lParam);
}

LRESULT CALLBACK WndProc( HWND hWnd, UINT messg, WPARAM wParam, LPARAM lParam )
{
	static HWND hPrecedente,hSuivante,hArreter,hActualiser,hDemarrage,hEnregistrer,hImprimer,hAller;
	static HWND hCadre1,hCadre2,hTadresse,hEtat;
	static int PageCounter=0;
	static int ObjCounter=0;
	char tampon[256];
	static BSTR titre;
	WCHAR url[256];

	switch(messg)
	{
		case WM_CREATE:
			// Création de tous les contrôles:
			hPrecedente=CreateWindow("BUTTON","Précédente",WS_CHILD | WS_VISIBLE,5,6,90,20,hWnd,0,0,0);
			hSuivante=CreateWindow("BUTTON","Suivante",WS_CHILD | WS_VISIBLE,105,6,90,20,hWnd,0,0,0);
			hArreter=CreateWindow("BUTTON","Arrêter",WS_CHILD | WS_VISIBLE,205,6,90,20,hWnd,0,0,0);
			hActualiser=CreateWindow("BUTTON","Actualiser",WS_CHILD | WS_VISIBLE,305,6,90,20,hWnd,0,0,0);
			hDemarrage=CreateWindow("BUTTON","Démarrage",WS_CHILD | WS_VISIBLE,405,6,90,20,hWnd,0,0,0);
			hEnregistrer=CreateWindow("BUTTON","Enregistrer",WS_CHILD | WS_VISIBLE,505,6,90,20,hWnd,0,0,0);
			hImprimer=CreateWindow("BUTTON","Imprimer",WS_CHILD | WS_VISIBLE,605,6,90,20,hWnd,0,0,0);
			hEtat=CreateStatusWindow(WS_CHILD|WS_VISIBLE,"Projet Silo C++ v:1.0.0.1",hWnd,6000);
			EnableWindow(hPrecedente,0);
			EnableWindow(hSuivante,0);
			EnableWindow(hArreter,0);
			EnableWindow(hActualiser,0);
			EnableWindow(hEnregistrer,0);
			EnableWindow(hImprimer,0);
			break;
		
		case WM_COMMAND: {
			if ((HWND)lParam == hPrecedente){ pIWeb->GoBack(); break; }
			if ((HWND)lParam == hSuivante){ pIWeb->GoForward(); break; }
			if ((HWND)lParam == hArreter){PageCounter = ObjCounter = 0;pIWeb->Stop();EnableWindow(hArreter, 0);EnableWindow(hActualiser, 1);SetWindowText(hEtat, "Arrêté");break;}
			if ((HWND)lParam == hActualiser){ pIWeb->Refresh2(0);	break; }
			if ((HWND)lParam == hDemarrage){ MultiByteToWideChar(CP_ACP, 0, "http://192.168.0.15", -1, url, 256);	pIWeb->Navigate(url, 0, 0, 0, 0); break; }
			if ((HWND)lParam == hEnregistrer){ pIWeb->ExecWB(OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT, 0, 0); break; }
			if ((HWND)lParam == hImprimer){ pIWeb->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0); break; }
		}
		case BEFORENAVIGATE2: {
			if ((IUnknown*)wParam==pIWeb)	PageCounter=ObjCounter=0;
			PageCounter++;
			EnableWindow(hArreter,1);
			EnableWindow(hActualiser,0);
			EnableWindow(hEnregistrer,0);
			EnableWindow(hImprimer,0);
		}break;
		case DOWNLOADBEGIN: {
			ObjCounter++;
			EnableWindow(hArreter,1);
			EnableWindow(hActualiser,0);
			EnableWindow(hEnregistrer,0);
			EnableWindow(hImprimer,0);
			SetWindowText(hEtat,"Navigation");		
		}break;
		case DOWNLOADCOMPLETE: {
			ObjCounter--;
			if (PageCounter==0 && ObjCounter==0){
				EnableWindow(hArreter,0);
				EnableWindow(hActualiser,1);
				EnableWindow(hEnregistrer,1);
				EnableWindow(hImprimer,1);
				SetWindowText(hEtat,"Terminé");
			}
		}break;
		case DOCUMENTCOMPLETE: {
			PageCounter--;
			if (PageCounter==0){	
				pIWeb->get_LocationName(&titre);
				WideCharToMultiByte(CP_ACP,0,titre,-1,tampon,256,0,0);
				lstrcat(tampon," - Projet Silo C++ v:1.0.0.1");
				SetWindowText(hWnd,tampon);
				EnableWindow(hArreter,0);
				EnableWindow(hActualiser,1);
				EnableWindow(hEnregistrer,1);
				EnableWindow(hImprimer,1);
				SetWindowText(hEtat,"Terminé");
			}
		}break;
		case NAVIGATECOMPLETE2: {
			if((IUnknown*)wParam==pIWeb){
				WideCharToMultiByte(CP_ACP,0,(LPCWSTR)lParam,-1,tampon,256,0,0);
				SetWindowText(hAdresse,tampon);
				pIWeb->get_LocationName(&titre);
				WideCharToMultiByte(CP_ACP,0,titre,-1,tampon,256,0,0);
				lstrcat(tampon," - Projet Silo C++ v:1.0.0.1");
				SetWindowText(hWnd,tampon);
			}
		}break;
		case COMMANDSTATECHANGE: {
			if (wParam==2) EnableWindow(hPrecedente,(BOOL)lParam);
			if (wParam==1) EnableWindow(hSuivante,(BOOL)lParam);
		}break;
		case WM_SIZE:{
			RECT rc;
			MoveWindow(hConteneur,5,25,LOWORD(lParam)-8, HIWORD(lParam)-100,1);
			GetClientRect(hConteneur,&rc);
			MoveWindow(hEtat,0,rc.bottom-32,rc.right, 25,1);
		}return 0;
		case WM_CLOSE: {
			SysFreeString(titre);
			DestroyWindow(hWnd);
		}break;
		case WM_DESTROY: {PostQuitMessage( 0 );	}break;
		default:return( DefWindowProc( hWnd, messg, wParam, lParam ) );
	}
	return 0;
}
int WINAPI WinMain( HINSTANCE hInst,HINSTANCE hPreInst,LPSTR lpszCmdLine, int nCmdShow ){
	WNDCLASS wc;
	char NomClasse[]   = "Projet_Silo_cpp";
	wc.lpszClassName 	= NomClasse;
	wc.hInstance 		= hInst;
	wc.lpfnWndProc		= WndProc;
	wc.hCursor			= LoadCursor( NULL, IDC_ARROW );
	wc.hIcon			= LoadIcon( wc.hInstance,(LPCSTR) 101 );
	wc.lpszMenuName	    = NULL;
	wc.hbrBackground	= GetSysColorBrush(COLOR_WINDOW);
	wc.style			= 0;
	wc.cbClsExtra		= 0;
	wc.cbWndExtra		= 0;
	if (!RegisterClass(&wc)) return 0;
	RECT rc;
	GetClientRect(GetDesktopWindow(), &rc);
	HWND hWnd = CreateWindow( NomClasse,"Projet Silo C++ v:1.0.0.1", WS_POPUP |WS_BORDER | WS_SYSMENU|WS_CAPTION,rc.left,rc.top,rc.right,rc.bottom, 0, 0, hInst,0);
	ShowWindow(hWnd, nCmdShow );
	UpdateWindow( hWnd );
	HINSTANCE hDLL = LoadLibrary("atl.dll");
	typedef HRESULT (WINAPI *PAttachControl)(IUnknown*, HWND,IUnknown**);
	PAttachControl AtlAxAttachControl = (PAttachControl) GetProcAddress(hDLL, "AtlAxAttachControl");
	RECT rect;
	GetClientRect(hWnd,&rect);
	hConteneur=CreateWindowEx(WS_EX_CLIENTEDGE,"BUTTON","",WS_CHILD | WS_VISIBLE,5,32,rect.right-8,rect.bottom-67,hWnd,0,0,0);
	CoInitialize(0);
	CoCreateInstance(CLSID_WebBrowser,0,CLSCTX_ALL,IID_IWebBrowser2,(void**)&pIWeb);
	AtlAxAttachControl(pIWeb,hConteneur,NULL);
	IConnectionPointContainer* pCPContainer;
    pIWeb->QueryInterface(IID_IConnectionPointContainer,(void**)&pCPContainer);
	IConnectionPoint *pConnectionPoint;
	pCPContainer->FindConnectionPoint(DIID_DWebBrowserEvents2, &pConnectionPoint);
	Evenem *pEvnm;
	pEvnm= new Evenem(hWnd);
	DWORD dwCookie = 0;
	pConnectionPoint->Advise(pEvnm, &dwCookie);
	if (pCPContainer) pCPContainer->Release();
	WCHAR url[256];
	MultiByteToWideChar (CP_ACP, 0,"http://192.168.0.15", -1, url, 256);
	pIWeb->Navigate(url,0,0,0,0);
	MSG Msg;
	while( GetMessage( &Msg, 0, 0, 0 ) ){
		TranslateMessage( &Msg );
		DispatchMessage( &Msg );
	}
	pConnectionPoint->Unadvise(dwCookie);
   pConnectionPoint->Release();
	delete pEvnm;
	pIWeb->Release();
    CoUninitialize();		
	FreeLibrary(hDLL);
	return( Msg.wParam);
}
