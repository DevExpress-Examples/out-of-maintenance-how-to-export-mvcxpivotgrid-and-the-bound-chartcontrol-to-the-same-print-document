using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Collections;
using DevExpressMvcApplication1.Models;
using DevExpress.Web.Mvc;
using DevExpress.XtraPivotGrid;
using System.IO;
using DevExpress.XtraPrinting;

using DevExpress.XtraCharts.Web;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraPrintingLinks;
using DevExpress.XtraCharts;
using System.Diagnostics;

namespace DevExpressMvcApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View(GetSalesPersonData());
        }

        public ActionResult PivotGridPartial()
        {
            ViewBag.PivotSettings = GetPivotGridSettings();
            return PartialView("PivotGridPartial", GetSalesPersonData());
        }
        public ActionResult ChartPartial()
        {
            ViewBag.ChartSettings = GetChartSettings();
            var chartModel = PivotGridExtension.GetDataObject(GetPivotGridSettings(), GetSalesPersonData());
            return PartialView("ChartPartial", chartModel);
        }

        public ActionResult ExportPivotChart()
        {
            PrintingSystem ps = new PrintingSystem();
            
            PrintableComponentLink pclPivot = new PrintableComponentLink(ps);
            pclPivot.Component = PivotGridExtension.CreatePrintableObject(GetPivotGridSettings(), GetSalesPersonData());

            PrintableComponentLink pclChart = new PrintableComponentLink(ps);
            using (MemoryStream chartLayout = new MemoryStream())
            {
                GetChartSettings().SaveToStream(chartLayout);
                MVCxChartControl chart = new MVCxChartControl();
                chart.LoadFromStream(chartLayout);
                chart.DataSource = PivotGridExtension.GetDataObject(GetPivotGridSettings(), GetSalesPersonData(), true);
                chart.Width = 900;
                chart.Height = 400;
                chart.DataBind();
                pclChart.Component = ((IChartContainer)chart).Chart;
            }

            CompositeLink compositeLink = new CompositeLink(ps);            
            compositeLink.Links.AddRange(new object[] { pclPivot, pclChart });
            compositeLink.CreateDocument();

            MemoryStream stream = new MemoryStream();
            compositeLink.PrintingSystem.ExportToXls(stream);
            stream.Position = 0;
            ps.Dispose();
            return File(stream, "application/ms-excel", "PivotWithChart.xls");
        }

        private ChartControlSettings GetChartSettings()
        {
            ChartControlSettings settings = new ChartControlSettings();
            settings.Name = "webChart";
            settings.CallbackRouteValues = new { Controller = "Home", Action = "ChartPartial" };
            settings.EnableClientSideAPI = true;
            settings.Legend.MaxHorizontalPercentage = 30;            
            settings.Width = System.Web.UI.WebControls.Unit.Pixel(830);
            settings.Height = System.Web.UI.WebControls.Unit.Pixel(300);
            settings.ClientSideEvents.BeginCallback = "OnBeginChartCallback";

            settings.SeriesDataMember = "Series";
            settings.SeriesTemplate.ChangeView(DevExpress.XtraCharts.ViewType.StackedBar);
            settings.SeriesTemplate.ArgumentDataMember = "Arguments";
            settings.SeriesTemplate.ValueDataMembers[0] = "Values";
            return settings;

        }



        public static System.Data.DataTable GetSalesPersonData()
        {
            return NwindModel.GetInvoices();
        }

        static PivotGridSettings GetPivotGridSettings()
        {
            PivotGridSettings settings = new PivotGridSettings();
            settings.Name = "pivotGrid";
            settings.CallbackRouteValues = new { Controller = "Home", Action = "PivotGridPartial" };
            settings.OptionsData.DataProcessingEngine = PivotDataProcessingEngine.Optimized;
            settings.OptionsView.HorizontalScrollBarMode = DevExpress.Web.ScrollBarMode.Auto;
            settings.OptionsChartDataSource.ProvideDataByColumns = false;
            settings.Width = new System.Web.UI.WebControls.Unit(90, System.Web.UI.WebControls.UnitType.Percentage);

            settings.Groups.Add("Order Date");
            settings.Fields.AddDataSourceColumn("Country", PivotArea.FilterArea);
            settings.Fields.AddDataSourceColumn("City", PivotArea.FilterArea);
            settings.Fields.Add(field =>
            {
                field.Area = PivotArea.ColumnArea;
                field.DataBinding = new DataSourceColumnBinding("OrderDate",  PivotGroupInterval.DateYear);
                field.Caption = "Year";
                field.GroupIndex = 0;
            });
            settings.Fields.Add(field =>
            {
                field.Area = PivotArea.ColumnArea;
                field.DataBinding = new DataSourceColumnBinding("OrderDate", PivotGroupInterval.DateMonth);
                field.Caption = "Month";
                field.GroupIndex = 0;
                field.InnerGroupIndex = 1;
            });

            settings.Fields.Add(field =>
            {
                field.Area = PivotArea.DataArea;
                field.DataBinding = new DataSourceColumnBinding("Quantity");
                field.Visible = false;
            });
            settings.Fields.Add(field =>
            {
                field.Area = PivotArea.DataArea;
                field.DataBinding = new DataSourceColumnBinding("UnitPrice");
                field.Visible = false;
            });
            settings.Fields.Add(field =>
            {
                field.Area = PivotArea.DataArea;
                field.DataBinding = new ExpressionDataBinding("[UnitPrice]*[Quantity]");
                field.Visible = true;
            });
            settings.OptionsData.AutoExpandGroups = DevExpress.Utils.DefaultBoolean.False;
            settings.Fields.AddDataSourceColumn("ProductName", PivotArea.RowArea);
            settings.ClientSideEvents.BeginCallback = "OnBeforePivotGridCallback";
            settings.ClientSideEvents.EndCallback = "UpdateChart";

            return settings;
        }

    }
}
