using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace COM.HPE
{
    public partial class BHO : IObjectWithSite
    {
        private string FindXPath(IHTMLElement ele)
        {
            StringBuilder builder = new StringBuilder();
            while (ele != null)
            {
                int index = FindElementIndex(ele);
                if (ele.id != null)
                {
                   
                    if (index == 1)
                    {
                        builder.Insert(0, "//*[@id='" + ele.id  + "']");
                    }
                    else
                    {
                        builder.Insert(0, "//*[@id='" + ele.id + "'][" + index + "]");
                    }
                   
                    return builder.ToString();
                }
                else
                {
                    if (index == 1)
                    {
                        builder.Insert(0, "/" + ele.tagName );
                    }
                    else
                    {
                        builder.Insert(0, "/" + ele.tagName  + "[" + index + "]");
                    }
                }
                ele = ele.parentElement;
            }
            
            return builder.ToString();
        }
        private string FindFullXPath(IHTMLElement ele)
        {
            StringBuilder builder = new StringBuilder();
            while (ele != null)
            {
                int index =  FindElementIndex(ele); 

                if (index == 1)
                {
                    builder.Insert(0, "/" + ele.tagName );
                }
                else
                {
                    builder.Insert(0, "/" + ele.tagName  + "[" + index + "]");
                }

                ele = ele.parentElement;
            }
            return builder.ToString();
        }
        private int FindElementIndex(IHTMLElement element)
        {
            IHTMLElement parent = element.parentElement;
            if (parent == null)
            {
                return 1;
            }

            int index = 1;
            foreach (IHTMLElement ele in (IHTMLElementCollection)parent.children)
            { 
               
                if (ele is IHTMLElement && ele.tagName == element.tagName && ele.className == element.className)
                {
                       if (ele == element)
                    { 
                        return index;
                    }
                
                    index++; 
                }
            }
            
            return index;
        }
        private string FindCssPath(IHTMLElement ele)
        {
             StringBuilder CssPath = new StringBuilder();
             while (ele != null)
             {
                 int index =  FindElementIndex(ele); 
                 if (ele.id != null)
                 {
                     CssPath.Insert(0, '#' + ele.id+ " > ");
                     break;
                 }
                 else
                 {
                     if (index == 1)
                         CssPath.Insert(0, ele.tagName + " > ");
                     else
                     {
                  //       CssPath.Insert(0, ele.tagName + ":nth-child(" + index + ")" + " > ");
                         CssPath.Insert(0, ele.tagName + " > ");
                       //  System.Windows.Forms.MessageBox.Show();
                     }
                     ele = ele.parentElement;
                 }
             }
           
            return CssPath.ToString().Remove(CssPath.Length-2);

        }

    }
}
