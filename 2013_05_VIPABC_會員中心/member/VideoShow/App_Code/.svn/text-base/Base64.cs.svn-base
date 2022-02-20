using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

/// <summary>
/// Base64 的摘要描述
/// </summary>
public class Base64
{
	public Base64()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
     public string EncodeTo64(string toEncode)
    {

        byte[] toEncodeAsBytes

              = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode);

        string returnValue

              = System.Convert.ToBase64String(toEncodeAsBytes);

        return returnValue;

    }
      public string DecodeFrom64(string encodedData)
     {

         try
         {
             byte[] encodedDataAsBytes

                 = System.Convert.FromBase64String(encodedData);

             string returnValue =

                System.Text.ASCIIEncoding.ASCII.GetString(encodedDataAsBytes);

             return returnValue;
         }
         catch
         {
             return "";
         }

     }


}
