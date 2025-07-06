# AllAutoOfficePDF2

Office文書（Word、Excel、PowerPoint）をPDFに変換し、結合するWPFアプリケーションです。

## 機能

- Office文書のPDF変換
- PDF結合
- プロジェクト管理
- ページ番号追加
- ファイル順序変更

## プロジェクト構造

```
AllAutoOfficePDF2/
├── Models/                    # データモデル
│   ├── ProjectData.cs        # プロジェクトデータ
│   ├── FileItemData.cs       # ファイルアイテムデータ
│   └── FileItem.cs           # ファイルアイテム
├── Views/                     # ビュー
│   ├── MainWindow.xaml       # メインウィンドウ
│   ├── MainWindow.xaml.cs    # メインウィンドウコードビハインド
│   ├── ProjectEditDialog.xaml # プロジェクト編集ダイアログ
│   └── ProjectEditDialog.xaml.cs # プロジェクト編集ダイアログコードビハインド
├── Services/                  # サービス
│   ├── ProjectManager.cs     # プロジェクト管理
│   ├── PdfConversionService.cs # PDF変換サービス
│   ├── PdfMergeService.cs    # PDF結合サービス
│   └── FileManagementService.cs # ファイル管理サービス
├── ViewModels/               # ビューモデル（将来の拡張用）
├── Converters/               # 値コンバーター（将来の拡張用）
├── Controls/                 # カスタムコントロール（将来の拡張用）
├── App.xaml                  # アプリケーション定義
├── App.xaml.cs               # アプリケーションコードビハインド
└── AssemblyInfo.cs           # アセンブリ情報
```

## 技術スタック

- .NET 6.0
- WPF (Windows Presentation Foundation)
- Microsoft Office Interop
- iTextSharp

## 依存関係

- Microsoft.Office.Interop.Word
- Microsoft.Office.Interop.Excel
- Microsoft.Office.Interop.PowerPoint
- iTextSharp
- System.Text.Json

## 使用方法

1. プロジェクトを作成または選択
2. 対象フォルダを選択
3. ファイルを読み込み
4. 必要に応じてファイル順序を変更
5. PDF変換を実行
6. PDF結合を実行

## 設計原則

- **責任分離**: ModelとServiceに分離
- **保守性**: 各機能を独立したクラスに分離
- **拡張性**: 将来の機能追加を考慮した構造
- **可読性**: 適切なコメントと名前付け

## 今後の拡張可能性

- ViewModels: MVVMパターンの本格実装
- Converters: データバインディングの値変換
- Controls: 再利用可能なカスタムコントロール
- 設定管理: アプリケーション設定の管理
- ログ機能: エラーログや操作ログの記録