object FormSendEmail: TFormSendEmail
  Left = 0
  Top = 0
  Caption = 'Enviar Email'
  ClientHeight = 166
  ClientWidth = 414
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Button2: TButton
    Left = 158
    Top = 135
    Width = 75
    Height = 25
    Caption = 'Enviar'
    TabOrder = 2
    OnClick = Button2Click
  end
  object edtRemetente: TLabeledEdit
    Left = 8
    Top = 24
    Width = 398
    Height = 21
    EditLabel.Width = 17
    EditLabel.Height = 13
    EditLabel.Caption = 'De:'
    TabOrder = 0
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 51
    Width = 398
    Height = 78
    Caption = 
      'Caso queira enviar para mais de um email utiliza ponto-e-virgula' +
      ' para separar'
    TabOrder = 1
    object edtDestinatario: TLabeledEdit
      Left = 3
      Top = 40
      Width = 392
      Height = 21
      EditLabel.Width = 26
      EditLabel.Height = 13
      EditLabel.Caption = 'Para:'
      TabOrder = 0
    end
  end
  object qryConfigEmail: TFDQuery
    Connection = DM.FDConnection1
    SQL.Strings = (
      'select * from espConfiguracaoEmail WHERE modulo :modulo')
    Left = 344
    Top = 107
    ParamData = <
      item
        Name = 'MODULO'
        DataType = ftString
        ParamType = ptInput
        Size = 250
      end>
  end
end
