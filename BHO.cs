using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using SHDocVw;
using mshtml;
using Microsoft.Win32;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels.Ipc;
using System.ServiceModel;
using System.Data;
using System.Globalization;
using System.Net;
using System.Net.Sockets; 
using System.Xml;
namespace COM.HPE
{
    [
     ComVisible(true),
     Guid("8a194578-81ea-4850-9911-13ba2d71efbd"),
     ClassInterface(ClassInterfaceType.None)
     ]
    [ServiceBehavior(ConcurrencyMode = ConcurrencyMode.Reentrant, InstanceContextMode = InstanceContextMode.Single)]
    public partial class BHO : IObjectWithSite
    {
        WebBrowser webBrowser;
        HTMLDocument document;
        System.Windows.Forms.Form f = new System.Windows.Forms.Form();
        System.Windows.Forms.Label l = new System.Windows.Forms.Label();

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            document = (HTMLDocument)webBrowser.Document;
            HTMLDocument doc = this.document as HTMLDocument;

            bindHandler(doc);
          
            object j;
            
            for (int i = 0; i < doc.parentWindow.frames.length; i++)
            {
                j = i;
                IHTMLWindow2 frame = doc.frames.item(ref j) as IHTMLWindow2;
                
                IHTMLDocument idoc = CodecentrixSample.CrossFrameIE.GetDocumentFromWindow(frame);

                bindIFrameHandler(idoc);

                findIFrame(frame);
            }
      
            insertScript();
        }
        private void findIFrame(IHTMLWindow2 frame)
        {
            
            if (frame.frames.length == 0)
                return;

            object j;
            for (int i = 0; i < frame.frames.length; i++)
            {
                j = i;
                IHTMLWindow2 childframe = frame.frames.item(ref j) as IHTMLWindow2;
                IHTMLDocument idoc = CodecentrixSample.CrossFrameIE.GetDocumentFromWindow(childframe);

                bindIFrameHandler(idoc);

                findIFrame(childframe);
            }
        }
        private void bindHandler(HTMLDocument doc)
        {
            DHTMLEventHandler onClick = new DHTMLEventHandler(doc);
            DHTMLEventHandler onMouseOver = new DHTMLEventHandler(doc);
            DHTMLEventHandler onMouseOut = new DHTMLEventHandler(doc);

            onClick.Handler += new DHTMLEvent(this.onClick);
            onMouseOver.Handler += new DHTMLEvent(this.OnMouseOver);
            onMouseOut.Handler += new DHTMLEvent(this.OnMouseOut);

            ((mshtml.DispHTMLDocument)doc).onmouseup = onClick;
            ((mshtml.DispHTMLDocument)doc).onmouseover = onMouseOver;
            ((mshtml.DispHTMLDocument)doc).onmouseout = onMouseOut;
        }
        private void bindIFrameHandler(IHTMLDocument idoc)
        {
            DHTMLEventHandler onIframeClick = new DHTMLEventHandler(idoc as HTMLDocument);
            DHTMLEventHandler onIframeMouseOver = new DHTMLEventHandler(idoc as HTMLDocument);
            DHTMLEventHandler onIframeMouseOut = new DHTMLEventHandler(idoc as HTMLDocument);

            onIframeClick.Handler += new DHTMLEvent(this.onIframeClick);
            onIframeMouseOver.Handler += new DHTMLEvent(this.OnMouseOver);
            onIframeMouseOut.Handler += new DHTMLEvent(this.OnMouseOut);

            ((mshtml.DispHTMLDocument)idoc).onmouseup = onIframeClick;
            ((mshtml.DispHTMLDocument)idoc).onmouseover = onIframeMouseOver;
            ((mshtml.DispHTMLDocument)idoc).onmouseout = onIframeMouseOut;
        }
        private void insertScript()
        {
            var htmlDoc = (IHTMLDocument3)webBrowser.Document;

            HTMLHeadElement head = htmlDoc.getElementsByTagName("head").Cast<HTMLHeadElement>().First();

            var script = (IHTMLScriptElement)((IHTMLDocument2)htmlDoc).createElement("script");

            string javaScriptText = @"  
		        var num; 
	            var shadow; 
	            var IEborder; 
	            var onOff = true; 
	            var status = false; 
	            var note = []; 
	            var boxShadowNote = []; 
	            var borderNote = []; 
	            var borderNum = 0; 
	            var borderNotes = []; 
	            var browser = navigator.appName; 
	            var b_version = navigator.appVersion; 
	            var version = b_version.split(';'); 
	            var trim_version = version[1].replace(/[ ]/g,'');	
 
	            document.body.onmouseover = function(el){ 
		            var el = el || window.event; 
		            if(browser == 'Microsoft Internet Explorer' && (trim_version == 'MSIE7.0' || trim_version == 'MSIE8.0')){
		                status = true; 
			            var borderWidth = el.srcElement.currentStyle.borderWidth; 
			            var borderColor = el.srcElement.currentStyle.borderColor; 
			            var borderStyle = el.srcElement.currentStyle.borderStyle; 
		            }else{ 
			            shadow = el.srcElement.style.boxShadow; 
		            }
		
		            document.body.onclick = function(){ 
			            onOff = false; 
			            if(status){ 
				            borderNote.push(borderWidth); 
				            borderNote.push(borderColor); 
				            borderNote.push(borderStyle); 
				            console.log(borderNote.length) 
				            console.log(borderNote) 
			            }else{ 
				            boxShadowNote.push(shadow);		 
			            }
			
			            ReadXMl();
			
			            el.srcElement.onmouseout = function(){ 
				            if(status){ 
					            this.style.border = '1px solid red'; 
				            }else{ 
					            this.style.boxShadow = '0 0 0 1px red'; 
				            } 
			            }
		            };
			
		            if(status){ 
			            el.srcElement.style.border = '1px solid #0CC'; 
		            }else{ 
			            var shadowStatus = el.srcElement.style.boxShadow; 
			            el.srcElement.style.boxShadow = '0 0 12px #0CC';
		            }
		
		
		            el.srcElement.onmouseout = function(){ 
			            if(status){ 
				            if(borderStyle == 'none' && borderWidth == 'medium'){ 
					            this.style.border = 'none'; 
				            }else{ 
					            this.style.borderWidth = borderWidth; 
					            this.style.borderStyle = borderStyle; 
					            this.style.borderColor = borderColor; 
				            } 
			            }else{ 
				            if(onOff){ 
					            this.style.boxShadow = shadowStatus; 
				            } 
			            } 
			            onOff = true;; 
		            }; 
	            } 
	
	            function ReadXMl(){ 
		            var timeout = setTimeout(function(){ 
		                var MarkPoint = document.querySelectorAll('.MarkPoint'); 
			                if(MarkPoint.length > 0){ 
	                            for (var i = MarkPoint.length - 1; i >= 0 ; i--) { 
	                                document.body.removeChild(MarkPoint[i]); 
			                    }
			                }; 
			
			            for (var i = 0; i < borderNote.length; i++) { 
				            borderNotes.push(borderNote[i].replace(/medium/g,'0')); 
			            }
			
		                for (var i = 0; i < note.length; i++) { 
		    	            if(status){ 
		    			            note[i].style.borderWidth = borderNotes[0 + borderNum] ; 
		    			            borderNum++; 
		    			            note[i].style.borderColor = borderNotes[0 + borderNum]; 
		    			            borderNum++; 
		    			            note[i].style.borderStyle = borderNotes[0 + borderNum]; 
		    			            borderNum++; 
		    	            }else{ 
		  			            note[i].style.boxShadow = boxShadowNote[i]; 
		    	            } 
		                };
		    
		                note = []; 
		                borderNum = 0; 
		 	            Read(); 
			            },200); 
	            };
	
		            function Read(){ 
			            var xmlDoc = new ActiveXObject('MSXML2.DOMDocument'); 
			            xmlDoc.load('c:\\Data\\Read.xml'); 
			            var Nodes = xmlDoc.getElementsByTagName('Nodes');
	
			            for (var i = 0; i < Nodes.length ; i++) {				
				            if(window.location.href == Nodes[i].firstChild.lastChild.text){					
					            for (var j = 0; j < Nodes[i].childNodes.length ; j++) {						
						            num = Nodes[i].childNodes[j].firstChild.text;						
						            ControlElement(Nodes[i].childNodes[j].firstChild.nextSibling.text)
					            };
				            }
			            }
		            }
		
		
		            function ControlElement(CssPath){ 
			            var Node = document.querySelector(CssPath); 
			            note.push(Node); 
			            if(browser == 'Microsoft Internet Explorer' && (trim_version == 'MSIE7.0'||trim_version == 'MSIE8.0')){ 
				            Node.style.border = '1px solid red'; 
			            }else{ 
				            Node.style.boxShadow = '0 0 0 1px red'; 
			            }			
			            Print(Node);
		            }
	
		            function Print(Node){			
			            var insert = document.createElement('div');	
			            var mark = document.createElement('p');			
			            mark.innerText = num;			
			            mark.style.width = '20px';			
			            mark.style.height = '20px';			
			            mark.style.fontSize = '0.9em';			
			            mark.style.lineHeight = '20px';			
			            mark.style.margin = '0 auto';			
			            mark.style.verticalAlign = 'middle';			
			            mark.style.color = 'white';			
			            mark.style.textAlign = 'center';			
			            insert.setAttribute('class','MarkPoint');			
			            insert.style.position = 'absolute';			
			            insert.style.zIndex = '1000';			
			            insert.style.width = '20px';			
			            insert.style.height = '20px';			
			            insert.style.borderRadius = '10px';			
			            insert.style.lineHeight = '20px';			
			            insert.style.background = 'red';			
			            insert.style.top = Node.getBoundingClientRect().top + document.documentElement.scrollTop + 'px';			
			            insert.style.left = Node.getBoundingClientRect().left + document.documentElement.scrollLeft  + 'px';						
			            insert.appendChild(mark);			
			            document.body.appendChild(insert); 
		            } 
                    ";
            script.text = javaScriptText;
            head.appendChild((IHTMLDOMNode)script);  
        }
        public int SetSite(object site)
        {
            try
            {
                if (site != null)
                {
                    webBrowser = (WebBrowser)site;
                    webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete); 
                }
                else
                {
                    webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                    webBrowser = null;
                } 
            }
            catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
            return 0;
        }
        void MainInfo(string msg)
        {
            byte[] data = new byte[1024];//定义一个数组用来做数据的缓冲区
 
            IPEndPoint ipep = new IPEndPoint(IPAddress.Parse("127.0.0.1"), 9050);
            Socket server = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp); 

            server.SendTo(Encoding.UTF8.GetBytes(msg), ipep);//将数据发送到指定的终结点Remote
            
            server.Close();
        }
   
        #region Handler
        public delegate void DHTMLEvent(IHTMLEventObj e);
        [ComVisible(true)]
        public class DHTMLEventHandler
        {
            public DHTMLEvent Handler;

            HTMLDocument Document;

            public DHTMLEventHandler(HTMLDocument doc)
            {
                this.Document = doc;
            }

            [DispId(0)]

            public void Call()
            {
                Handler(Document.parentWindow.@event);
            }
        }
        #endregion
     

        #region MainBHO
        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }

        public static string BHOKEYNAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);
            if (registryKey == null)
            {
                registryKey = Registry.LocalMachine.CreateSubKey(BHOKEYNAME);
            }

            string guid = type.GUID.ToString("B");
            RegistryKey bhoKey = registryKey.OpenSubKey(guid, true);
            if (bhoKey == null)
            {
                bhoKey = registryKey.CreateSubKey(guid);
            }
            // NoExplorer: dword = 1 prevents the BHO to be loaded by Explorer.exe
            bhoKey.SetValue("NoExplorer", 1);
            bhoKey.Close();

            registryKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }
        #endregion
    }  
}
