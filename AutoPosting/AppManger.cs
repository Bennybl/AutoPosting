using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace AutoPosting
{
    class AppManger
    {
        private String m_SpreadsheetId = "1Mh0Gdmm81cLc_8Djjit-fwQKxqrJLosWoUsBC5fG8IE";

        private string m_SheetRange = "sheet2!A:D";

        private ValueRange m_Sheet;

        private string m_FacebookAccessToken;

        private string m_PageId;

        private ConfessionPost m_ConfessionPost;

        private FacebookApi m_FacebookApi;

        private GoogleSheetAPI m_SheetAPI;

        public AppManger()
        {
            m_SheetAPI = new GoogleSheetAPI(m_SpreadsheetId, m_SheetRange);
            m_Sheet = m_SheetAPI.GetSheet();
            m_FacebookAccessToken = m_Sheet.Values[1][2].ToString();
            m_PageId = m_Sheet.Values[1][3].ToString();
            m_FacebookApi = new FacebookApi(m_PageId, m_FacebookAccessToken);
            
            setConfessionObj();
            postToFacebook();
        }

        private void setConfessionObj()
        {
            m_ConfessionPost = new ConfessionPost(
                 m_Sheet.Values[2][0].ToString(),
                 Int32.Parse(m_Sheet.Values[1][1].ToString())
                );
        }

        //Needs to increment the confNum
        private void postToFacebook()
        {
            var post = m_FacebookApi.PublishMessage(m_ConfessionPost.ToString());

            try{
                _ = post.Result;
            }catch(Exception e)
            {
                m_ConfessionPost.IsPosted = false;
                m_ConfessionPost.exceptionInPosting = e;
            }
            if (m_ConfessionPost.IsPosted)
            {
                
                m_SheetAPI.deleteRow();
                m_SheetAPI.UpdateCell("sheet2!B2", new List<object>() { Int32.Parse(m_Sheet.Values[1][1].ToString()) + 1 });
            }
            updateConfToDataBase();

        }

        private void updateConfToDataBase()
        {
            string exeptionMessage = "";
            if(m_ConfessionPost.exceptionInPosting != null)
            {
                exeptionMessage = m_ConfessionPost.exceptionInPosting.Message;
            }

            var range = $"sheet4!{m_Sheet.Values[0][3].ToString()}:{m_Sheet.Values[0][3].ToString()}";
            List<Object> objList = new List<object>()
            {
                m_ConfessionPost.DateCreated,
                m_ConfessionPost.PostNumber,
                m_ConfessionPost.Post,
                m_ConfessionPost.IsPosted,
                exeptionMessage

            };
            m_SheetAPI.UpdateCell("sheet4!2:2", objList);
            m_SheetAPI.UpdateCell("sheet2!D1", new List<object>() { Int32.Parse(m_Sheet.Values[0][3].ToString()) + 1 });
            //ToDo
        }
    }
}
