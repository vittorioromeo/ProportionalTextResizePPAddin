using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ProportionalTextResizePPAddin
{
    public partial class ThisAddIn
    {
        public bool IsFeatureEnabled { get; set; } = true;

        private PowerPoint.Shape _trackedShape = null;
        private float _originalWidth;
        private float _originalHeight;

        private readonly Dictionary<int, float> _originalFontSizes = new Dictionary<int, float>();
        private readonly Dictionary<int, Tuple<Office.MsoTriState, float>> _originalParaSpacing = new Dictionary<int, Tuple<Office.MsoTriState, float>>();

        // Timer to actively check for resize changes on the selected shape.
        private Timer _resizeCheckTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Initialize the timer
            _resizeCheckTimer = new Timer();
            _resizeCheckTimer.Interval = 25; // Check every 25ms
            _resizeCheckTimer.Tick += ResizeCheckTimer_Tick;

            // Subscribe to the main PowerPoint event
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up resources
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
            _resizeCheckTimer.Tick -= ResizeCheckTimer_Tick;
            _resizeCheckTimer.Dispose();
        }

        private bool IsShiftKeyPressed()
        {
            return (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            // Stop monitoring any previously selected shape.
            _resizeCheckTimer.Stop();
            ResetTrackedShape();

            // If the feature is disabled, do nothing further.
            if (!IsFeatureEnabled) return;

            // Check if the new selection is a single shape we should monitor.
            if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && Sel.ShapeRange.Count == 1)
            {
                PowerPoint.Shape currentShape = Sel.ShapeRange[1];

                if (currentShape.HasTextFrame == Office.MsoTriState.msoTrue &&
                    currentShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    // A valid shape is selected. Record its state and start the timer.
                    UpdateTrackedShapeState(currentShape);
                    _resizeCheckTimer.Start();
                }
            }
        }

        private void ResizeCheckTimer_Tick(object sender, EventArgs e)
        {
            if (_trackedShape == null)
            {
                _resizeCheckTimer.Stop();
                return;
            }

            // If the feature is disabled, do nothing further.
            if (!IsFeatureEnabled) return;

            try
            {
                // Check if the dimensions have changed since our last check.
                if (_trackedShape.Width != _originalWidth || _trackedShape.Height != _originalHeight)
                {
                    if (IsShiftKeyPressed())
                    {
                        ProportionallyResizeTextInShape(_trackedShape);
                    }

                    UpdateTrackedShapeState(_trackedShape);
                }
            }
            catch (COMException)
            {
                // Shape was deleted. Stop monitoring.
                _resizeCheckTimer.Stop();
                ResetTrackedShape();
            }
        }

        private void ProportionallyResizeTextInShape(PowerPoint.Shape shape)
        {
            if (_originalWidth <= 0) return;

            float widthRatio = shape.Width / _originalWidth;

            // Scale Font Sizes
            var textRuns = shape.TextFrame.TextRange.Runs();
            for (int i = 1; i <= textRuns.Count + 1; i++)
            {
                if (_originalFontSizes.TryGetValue(i, out float originalFontSize))
                {
                    float newFontSize = originalFontSize * widthRatio;
                    textRuns.Runs(i).Font.Size = newFontSize;
                }
            }

            // Scale Paragraph Spacing
            var paragraphs = shape.TextFrame.TextRange.Paragraphs();
            for (int i = 1; i <= paragraphs.Count; i++)
            {
                if (_originalParaSpacing.TryGetValue(i, out var originalSpacing))
                {
                    var paraFormat = shape.TextFrame.TextRange.Paragraphs(i).ParagraphFormat;
                    var originalRule = originalSpacing.Item1;
                    var originalValue = originalSpacing.Item2;

                    paraFormat.LineRuleWithin = originalRule;

                    if (originalRule == Office.MsoTriState.msoFalse) // Spacing is in Points
                    {
                        paraFormat.SpaceWithin = originalValue * widthRatio;
                    }
                    else // Spacing is in Lines
                    {
                        paraFormat.SpaceWithin = originalValue;
                    }
                }
            }
        }

        private void UpdateTrackedShapeState(PowerPoint.Shape shape)
        {
            _trackedShape = shape;
            _originalWidth = shape.Width;
            _originalHeight = shape.Height;

            _originalFontSizes.Clear();
            var textRuns = shape.TextFrame.TextRange.Runs();
            for (int i = 1; i <= textRuns.Count; i++)
            {
                _originalFontSizes[i] = textRuns.Runs(i).Font.Size;
            }

            _originalFontSizes[textRuns.Count + 1] = (textRuns.Runs(textRuns.Count).Font.Size); // Workaround for last line...

            _originalParaSpacing.Clear();
            var paragraphs = shape.TextFrame.TextRange.Paragraphs();
            for (int i = 1; i <= paragraphs.Count; i++)
            {
                var paraFormat = shape.TextFrame.TextRange.Paragraphs(i).ParagraphFormat;
                _originalParaSpacing[i] = new Tuple<Office.MsoTriState, float>(paraFormat.LineRuleWithin, paraFormat.SpaceWithin);
            }
        }

        private void ResetTrackedShape()
        {
            _trackedShape = null;
            _originalFontSizes.Clear();
            _originalParaSpacing.Clear();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
    }
}
