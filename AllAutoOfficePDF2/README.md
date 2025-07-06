# AllAutoOfficePDF2

Office�����iWord�AExcel�APowerPoint�j��PDF�ɕϊ����A��������WPF�A�v���P�[�V�����ł��B

## �@�\

- Office������PDF�ϊ�
- PDF����
- �v���W�F�N�g�Ǘ�
- �y�[�W�ԍ��ǉ�
- �t�@�C�������ύX

## �v���W�F�N�g�\��

```
AllAutoOfficePDF2/
������ Models/                    # �f�[�^���f��
��   ������ ProjectData.cs        # �v���W�F�N�g�f�[�^
��   ������ FileItemData.cs       # �t�@�C���A�C�e���f�[�^
��   ������ FileItem.cs           # �t�@�C���A�C�e��
������ Views/                     # �r���[
��   ������ MainWindow.xaml       # ���C���E�B���h�E
��   ������ MainWindow.xaml.cs    # ���C���E�B���h�E�R�[�h�r�n�C���h
��   ������ ProjectEditDialog.xaml # �v���W�F�N�g�ҏW�_�C�A���O
��   ������ ProjectEditDialog.xaml.cs # �v���W�F�N�g�ҏW�_�C�A���O�R�[�h�r�n�C���h
������ Services/                  # �T�[�r�X
��   ������ ProjectManager.cs     # �v���W�F�N�g�Ǘ�
��   ������ PdfConversionService.cs # PDF�ϊ��T�[�r�X
��   ������ PdfMergeService.cs    # PDF�����T�[�r�X
��   ������ FileManagementService.cs # �t�@�C���Ǘ��T�[�r�X
������ ViewModels/               # �r���[���f���i�����̊g���p�j
������ Converters/               # �l�R���o�[�^�[�i�����̊g���p�j
������ Controls/                 # �J�X�^���R���g���[���i�����̊g���p�j
������ App.xaml                  # �A�v���P�[�V������`
������ App.xaml.cs               # �A�v���P�[�V�����R�[�h�r�n�C���h
������ AssemblyInfo.cs           # �A�Z���u�����
```

## �Z�p�X�^�b�N

- .NET 6.0
- WPF (Windows Presentation Foundation)
- Microsoft Office Interop
- iTextSharp

## �ˑ��֌W

- Microsoft.Office.Interop.Word
- Microsoft.Office.Interop.Excel
- Microsoft.Office.Interop.PowerPoint
- iTextSharp
- System.Text.Json

## �g�p���@

1. �v���W�F�N�g���쐬�܂��͑I��
2. �Ώۃt�H���_��I��
3. �t�@�C����ǂݍ���
4. �K�v�ɉ����ăt�@�C��������ύX
5. PDF�ϊ������s
6. PDF���������s

## �݌v����

- **�ӔC����**: Model��Service�ɕ���
- **�ێ琫**: �e�@�\��Ɨ������N���X�ɕ���
- **�g����**: �����̋@�\�ǉ����l�������\��
- **�ǐ�**: �K�؂ȃR�����g�Ɩ��O�t��

## ����̊g���\��

- ViewModels: MVVM�p�^�[���̖{�i����
- Converters: �f�[�^�o�C���f�B���O�̒l�ϊ�
- Controls: �ė��p�\�ȃJ�X�^���R���g���[��
- �ݒ�Ǘ�: �A�v���P�[�V�����ݒ�̊Ǘ�
- ���O�@�\: �G���[���O�⑀�샍�O�̋L�^