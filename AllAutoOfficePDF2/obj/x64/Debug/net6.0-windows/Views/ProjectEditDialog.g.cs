﻿#pragma checksum "..\..\..\..\..\Views\ProjectEditDialog.xaml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "C38D28CFF342D272B53AF5B42395C592359CA95F"
//------------------------------------------------------------------------------
// <auto-generated>
//     このコードはツールによって生成されました。
//     ランタイム バージョン:4.0.30319.42000
//
//     このファイルへの変更は、以下の状況下で不正な動作の原因になったり、
//     コードが再生成されるときに損失したりします。
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Controls.Ribbon;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.Integration;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace AllAutoOfficePDF2.Views {
    
    
    /// <summary>
    /// ProjectEditDialog
    /// </summary>
    public partial class ProjectEditDialog : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 35 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox txtProjectName;
        
        #line default
        #line hidden
        
        
        #line 46 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox cmbCategory;
        
        #line default
        #line hidden
        
        
        #line 68 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox txtFolderPath;
        
        #line default
        #line hidden
        
        
        #line 72 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBlock txtFolderDropHint;
        
        #line default
        #line hidden
        
        
        #line 85 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.CheckBox chkIncludeSubfolders;
        
        #line default
        #line hidden
        
        
        #line 92 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.CheckBox chkUseCustomPdfPath;
        
        #line default
        #line hidden
        
        
        #line 98 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Grid gridCustomPdfPath;
        
        #line default
        #line hidden
        
        
        #line 114 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox txtCustomPdfPath;
        
        #line default
        #line hidden
        
        
        #line 118 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBlock txtPdfDropHint;
        
        #line default
        #line hidden
        
        
        #line 151 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnOK;
        
        #line default
        #line hidden
        
        
        #line 153 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnCancel;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "9.0.6.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/AllAutoOfficePDF2;component/views/projecteditdialog.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "9.0.6.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            
            #line 8 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((AllAutoOfficePDF2.Views.ProjectEditDialog)(target)).Loaded += new System.Windows.RoutedEventHandler(this.Window_Loaded);
            
            #line default
            #line hidden
            return;
            case 2:
            this.txtProjectName = ((System.Windows.Controls.TextBox)(target));
            return;
            case 3:
            this.cmbCategory = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 4:
            
            #line 63 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).DragEnter += new System.Windows.DragEventHandler(this.FolderDropArea_DragEnter);
            
            #line default
            #line hidden
            
            #line 64 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).DragOver += new System.Windows.DragEventHandler(this.FolderDropArea_DragOver);
            
            #line default
            #line hidden
            
            #line 65 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).DragLeave += new System.Windows.DragEventHandler(this.FolderDropArea_DragLeave);
            
            #line default
            #line hidden
            
            #line 66 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).Drop += new System.Windows.DragEventHandler(this.FolderDropArea_Drop);
            
            #line default
            #line hidden
            return;
            case 5:
            this.txtFolderPath = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.txtFolderDropHint = ((System.Windows.Controls.TextBlock)(target));
            return;
            case 7:
            
            #line 80 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.BtnSelectFolder_Click);
            
            #line default
            #line hidden
            return;
            case 8:
            this.chkIncludeSubfolders = ((System.Windows.Controls.CheckBox)(target));
            
            #line 86 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            this.chkIncludeSubfolders.Checked += new System.Windows.RoutedEventHandler(this.ChkIncludeSubfolders_Checked);
            
            #line default
            #line hidden
            
            #line 87 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            this.chkIncludeSubfolders.Unchecked += new System.Windows.RoutedEventHandler(this.ChkIncludeSubfolders_Unchecked);
            
            #line default
            #line hidden
            return;
            case 9:
            this.chkUseCustomPdfPath = ((System.Windows.Controls.CheckBox)(target));
            
            #line 93 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            this.chkUseCustomPdfPath.Checked += new System.Windows.RoutedEventHandler(this.ChkUseCustomPdfPath_Checked);
            
            #line default
            #line hidden
            
            #line 94 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            this.chkUseCustomPdfPath.Unchecked += new System.Windows.RoutedEventHandler(this.ChkUseCustomPdfPath_Unchecked);
            
            #line default
            #line hidden
            return;
            case 10:
            this.gridCustomPdfPath = ((System.Windows.Controls.Grid)(target));
            return;
            case 11:
            
            #line 109 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).DragEnter += new System.Windows.DragEventHandler(this.PdfDropArea_DragEnter);
            
            #line default
            #line hidden
            
            #line 110 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).DragOver += new System.Windows.DragEventHandler(this.PdfDropArea_DragOver);
            
            #line default
            #line hidden
            
            #line 111 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).DragLeave += new System.Windows.DragEventHandler(this.PdfDropArea_DragLeave);
            
            #line default
            #line hidden
            
            #line 112 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Border)(target)).Drop += new System.Windows.DragEventHandler(this.PdfDropArea_Drop);
            
            #line default
            #line hidden
            return;
            case 12:
            this.txtCustomPdfPath = ((System.Windows.Controls.TextBox)(target));
            return;
            case 13:
            this.txtPdfDropHint = ((System.Windows.Controls.TextBlock)(target));
            return;
            case 14:
            
            #line 126 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.BtnSelectCustomPdfPath_Click);
            
            #line default
            #line hidden
            return;
            case 15:
            this.btnOK = ((System.Windows.Controls.Button)(target));
            
            #line 152 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            this.btnOK.Click += new System.Windows.RoutedEventHandler(this.BtnOK_Click);
            
            #line default
            #line hidden
            return;
            case 16:
            this.btnCancel = ((System.Windows.Controls.Button)(target));
            
            #line 154 "..\..\..\..\..\Views\ProjectEditDialog.xaml"
            this.btnCancel.Click += new System.Windows.RoutedEventHandler(this.BtnCancel_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

