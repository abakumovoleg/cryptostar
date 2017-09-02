using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;

namespace Cryptostar
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {        
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='Cryptostar'>
            <group id='group1' label='Cryptostar'>
              <button id='button1' image='bitcoin' size='large' label='Get Data' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }
        
        public void OnButtonPressed(IRibbonControl control)
        {
            try
            {
                var dataLoader = new DataLoader();

                var tickers = dataLoader.LoadTickers();

                var dataRender = new DataRender();

                dataRender.RenderData(tickers);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }

}
