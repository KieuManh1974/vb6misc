using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;


public partial class MasterPage : System.Web.UI.MasterPage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //put the active id on current menu
        string page = Request.RawUrl.Substring(1).Replace(".aspx", string.Empty);
        page = page.Replace("OrchardRetreat/", string.Empty);
        page = page.Substring(0, 1).ToUpper() + page.Substring(1);


        HtmlGenericControl activeMenu = null;

        try
        {
            activeMenu = (HtmlGenericControl)FindControl("mnu_" + page);

            if (activeMenu.HasControls()) //remove anchor element
            {
                activeMenu.InnerHtml = @"<a id=""active"">" + page + "</a>";
            }
        }
        catch
        {
            //page does not have mnu_ control
        }
            
    }
}
