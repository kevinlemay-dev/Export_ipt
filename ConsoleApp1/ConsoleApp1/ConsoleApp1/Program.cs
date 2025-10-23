using ConsoleApp1;
using Inventor;
using System;
using System.Runtime.InteropServices;
using WinFormsApp = System.Windows.Forms.Application;

namespace InventorIdwViews
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            WinFormsApp.EnableVisualStyles();
            WinFormsApp.SetCompatibleTextRenderingDefault(false);

            using (var form = new UserInputForm())
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    RunInventorAutomation(form.ModelPath);
                }
            }
        }

        public static void RunInventorAutomation(string modelPath)
        {
            Inventor.Application invApp = null;

            try
            {
                invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            }
            catch
            {
                Type invType = Type.GetTypeFromProgID("Inventor.Application");
                invApp = (Inventor.Application)Activator.CreateInstance(invType);
                invApp.Visible = true;
            }

            // ----------------------------
            // Open the model
            // ----------------------------
            _Document modelDoc = (_Document)invApp.Documents.Open(modelPath, false);
            Console.WriteLine("Opened: " + modelDoc.DisplayName);

            // ----------------------------
            // Create a new drawing
            // ----------------------------
            DrawingDocument drawDoc = (DrawingDocument)invApp.Documents.Add(
                DocumentTypeEnum.kDrawingDocumentObject,
                invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject));

            Sheet sheet = drawDoc.Sheets[1];

            // ----------------------------
            // Get model size
            // ----------------------------
            Box rangeBox = null;
            if (modelDoc is PartDocument partDoc)
                rangeBox = partDoc.ComponentDefinition.RangeBox;
            else if (modelDoc is AssemblyDocument asmDoc)
                rangeBox = asmDoc.ComponentDefinition.RangeBox;
            else
                throw new InvalidOperationException("The selected file is not a part or assembly.");

            double modelWidth = Math.Abs(rangeBox.MaxPoint.X - rangeBox.MinPoint.X);
            double modelHeight = Math.Abs(rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y);
            double modelDepth = Math.Abs(rangeBox.MaxPoint.Z - rangeBox.MinPoint.Z);

            double offset = Math.Max(modelWidth, Math.Max(modelHeight, modelDepth)) * 2;
            double scale = 1.0;

            double sheetWidth = sheet.Width;
            double sheetHeight = sheet.Height;

            // ----------------------------
            // Place views
            // ----------------------------
            Point2d basePoint = invApp.TransientGeometry.CreatePoint2d(sheetWidth / 2 + offset, sheetHeight / 2 + offset);
            DrawingView baseView = sheet.DrawingViews.AddBaseView(
                modelDoc, basePoint, scale,
                ViewOrientationTypeEnum.kFrontViewOrientation,
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            Point2d topPoint = invApp.TransientGeometry.CreatePoint2d(basePoint.X, basePoint.Y + offset);
            sheet.DrawingViews.AddProjectedView(baseView, topPoint,
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            Point2d rightPoint = invApp.TransientGeometry.CreatePoint2d(basePoint.X + offset, basePoint.Y);
            sheet.DrawingViews.AddProjectedView(baseView, rightPoint,
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            Point2d isoPoint = invApp.TransientGeometry.CreatePoint2d(basePoint.X + offset, basePoint.Y + offset);
            sheet.DrawingViews.AddProjectedView(baseView, isoPoint,
                DrawingViewStyleEnum.kShadedDrawingViewStyle);

            drawDoc.Update();

            // ----------------------------
            // Save the .IDW file
            // ----------------------------
            string idwPath = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(modelPath),
                System.IO.Path.GetFileNameWithoutExtension(modelPath) + ".idw");

            drawDoc.SaveAs(idwPath, true);
            Console.WriteLine("Saved IDW: " + idwPath);
            
            // ----------------------------
            // Export to DXF
            // ----------------------------
            string dxfPath = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(modelPath),
                System.IO.Path.GetFileNameWithoutExtension(modelPath) + ".dxf");

            TranslatorAddIn dxfAddIn = (TranslatorAddIn)invApp.ApplicationAddIns.ItemById[
                "{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}"]; // DXF Translator Add-In

            if (dxfAddIn != null)
            {
                if (!dxfAddIn.Activated)
                    dxfAddIn.Activate();

                TranslationContext context = invApp.TransientObjects.CreateTranslationContext();
                context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                NameValueMap options = invApp.TransientObjects.CreateNameValueMap();
                options.Add("Sheet_Range", PrintRangeEnum.kPrintAllSheets);
                options.Add("Export_Acad_Objects", true);
                options.Add("Export_Acad_RevisionTable", false);
                options.Add("Auto_Align_Views", true);
                options.Add("Simplify_Drawing", false);
                options.Add("Remove_Line_Weights", false);
                options.Add("PromptSaveAsResult", false);
                options.Add("SilentOperation", true);
                options.Add("ACADVer", "R2018");

                DataMedium dataMedium = invApp.TransientObjects.CreateDataMedium();
                dataMedium.FileName = dxfPath;

                if (dxfAddIn.HasSaveCopyAsOptions[drawDoc, context, options])
                {
                    dxfAddIn.SaveCopyAs(drawDoc, context, options, dataMedium);
                    Console.WriteLine("Exported DXF: " + dxfPath);
                }
                else
                {
                Console.WriteLine("DXF export options not supported for this document.");
                }
            }
            else
            {
                Console.WriteLine("DXF Translator Add-In not found!");
            }

            // // ----------------------------
            // // Export to DWG (fixed)
            // // ----------------------------
            // string dwgPath = System.IO.Path.Combine(
            //     System.IO.Path.GetDirectoryName(modelPath),
            //     System.IO.Path.GetFileNameWithoutExtension(modelPath) + ".dwg");

            // TranslatorAddIn dwgAddIn = (TranslatorAddIn)invApp.ApplicationAddIns.ItemById[
            //     "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"]; // DWG Translator Add-In

            // if (dwgAddIn != null)
            // {
            //     if (!dwgAddIn.Activated)
            //         dwgAddIn.Activate();

            //     TranslationContext context = invApp.TransientObjects.CreateTranslationContext();
            //     context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

            //     NameValueMap options = invApp.TransientObjects.CreateNameValueMap();

            //     // ✅ These options ensure actual 2D geometry is exported
            //     options.Add("Sheet_Range", PrintRangeEnum.kPrintAllSheets);
            //     options.Add("Export_Acad_Objects", true);
            //     options.Add("Export_Acad_RevisionTable", false);
            //     options.Add("Auto_Align_Views", true);
            //     options.Add("Simplify_Drawing", false);
            //     options.Add("Remove_Line_Weights", false);
            //     options.Add("PromptSaveAsResult", false);
            //     options.Add("SilentOperation", true);
            //     options.Add("ACADVer", "R2018");

            //     DataMedium dataMedium = invApp.TransientObjects.CreateDataMedium();
            //     dataMedium.FileName = dwgPath;

            //     // ✅ Check if translator supports SaveCopyAs
            //     if (dwgAddIn.HasSaveCopyAsOptions[drawDoc, context, options])
            //     {
            //         dwgAddIn.SaveCopyAs(drawDoc, context, options, dataMedium);
            //         Console.WriteLine("Exported DWG: " + dwgPath);
            //     }
            //     else
            //     {
            //         Console.WriteLine("DWG export options not supported for this document.");
            //     }
            // }
            // else
            // {
            //     Console.WriteLine("DWG Translator Add-In not found!");
            // }

            // ----------------------------
            // Cleanup: close documents
            // ----------------------------
            drawDoc.Close(true);
            modelDoc.Close(true);

            Console.WriteLine("✅ Done!");
        }
    }
}
