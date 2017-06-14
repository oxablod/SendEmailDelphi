unit USendEmail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, IdSMTP, IdSSLOpenSSL, IdMessage, IdText,
  IdAttachmentFile,
  IdExplicitTLSClientServerBase, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdMessageClient, IdSMTPBase, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Buttons, Vcl.ComCtrls, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack,
  IdSSL, Vcl.Menus, IdAntiFreezeBase, Vcl.IdAntiFreeze, Vcl.Samples.Gauges,
  UITypes, FireDac.Stan.Param, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client;

type
  TFormSendEmail = class(TForm)
    Button2: TButton;
    edtRemetente: TLabeledEdit;
    GroupBox1: TGroupBox;
    edtDestinatario: TLabeledEdit;
    qryConfigEmail: TFDQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormSendEmail: TFormSendEmail;

implementation

{$R *.dfm}

uses UAutorizacao, DataModule, UFuncoes;

procedure TFormSendEmail.Button2Click(Sender: TObject);
var
  // objetos necessários para o funcionamento
  IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL;
  IdSMTP: TIdSMTP;
  IdMessage: TIdMessage;
  CaminhoAnexo, Anexo2: string;
  x, texto,host,usuario,senha: string;
  i,porta: integer;
const
  empresas: array [1 .. 4] of string = ('6', '7', '8', '14');
begin
  for i := 1 to 5 do
  begin
    if FormPrincipal.cd_empresa.Text = empresas[i] then
      x := empresas[i];
  end;
  // instanciação dos objetos
  IdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  IdSMTP := TIdSMTP.Create(Self);
  IdMessage := TIdMessage.Create(Self);
  // ShowMessage('teste 1');
  try
    // Configuração do SSL
    IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
    IdSSLIOHandlerSocket.SSLOptions.Mode := sslmClient;

    // Configuração do SMTP
    qryConfigEmail.Close;
    qryConfigEmail.SQl.Clear;
    qryConfigEmail.SQl.Text := 'SELECT * FROM espConfiguracaoEmail WHERE modulo = :modulo';
    qryConfigEmail.ParamByName('modulo').Value := 'Autorizacao';
    qryConfigEmail.Open;

    porta := qryConfigEmail.FieldByName('porta').AsInteger;
    host := qryConfigEmail.FieldByName('host').AsString;
    usuario := qryConfigEmail.FieldByName('usuario').AsString;
    senha := qryConfigEmail.FieldByName('senha').AsString;

    IdSMTP.IOHandler := IdSSLIOHandlerSocket;
    IdSMTP.AuthType := satDefault;
    IdSMTP.Port := porta;
    IdSMTP.Host := host;
    IdSMTP.Username := usuario;
    IdSMTP.Password := senha;
    // ShowMessage('teste 2');
    // Tentativa de conexão e autenticação
    try
      IdSMTP.Connect;
      IdSMTP.Authenticate;
    except
      on E: Exception do
      begin
        MessageDlg('Erro na conexão e/ou autenticação contatar o TI: ' +
          E.Message, mtWarning, [mbOK], 0);
        Exit;
      end;
    end;

    if (FormPrincipal.rgTipo.ItemIndex = 0) then
    begin
      texto := 'Autorização de Devolução ' + FormPrincipal.nroAutorizacao.Text +
        ', responder o email para ' + edtRemetente.Text;
    end;
    if (FormPrincipal.rgTipo.ItemIndex = 1) then
    begin
      texto := 'Autorização de Garantia ' + FormPrincipal.nroAutorizacao.Text +
        ', responder o email para ' + edtRemetente.Text;
    end;

    // Configuração da mensagem
    IdMessage.From.Address := edtRemetente.Text;
    IdMessage.From.Name := FormPrincipal.NomeEmpresa.Text;
    IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
    IdMessage.Recipients.EMailAddresses := edtDestinatario.Text;
    IdMessage.Subject := FormPrincipal.FDQPDF.FieldByName('msg_auto').AsString;
    IdMessage.Body.Text := texto;

    try
      if (FormPrincipal.cd_empresa.Text = '1') then
      begin
        // Anexo da mensagem (opcional)
        if (FormPrincipal.rgTipo.ItemIndex = 0) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\1\Devolucao\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
        end;
        if (FormPrincipal.rgTipo.ItemIndex = 1) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\1\Garantia\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          Anexo2 := '\\192.168.0.2\Diversos\AUTORIZACOES\1\Garantia\' +
            'Ficha_Garantia.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
          TIdAttachmentFile.Create(IdMessage.MessageParts, Anexo2);
        end;
      end;
      if (FormPrincipal.cd_empresa.Text = x) then
      begin
        // Anexo da mensagem (opcional)
        if (FormPrincipal.rgTipo.ItemIndex = 0) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\' + x +
            '\Devolucao\' + 'Autorização - ' + FormPrincipal.nroAutorizacao.Text
            + ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
        end;
        if (FormPrincipal.rgTipo.ItemIndex = 1) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\' + x +
            '\Garantia\' + 'Autorização - ' + FormPrincipal.nroAutorizacao.Text
            + ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          Anexo2 := '\\192.168.0.2\Diversos\AUTORIZACOES\' + x + '\Garantia\' +
            'Ficha_Garantia.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
          TIdAttachmentFile.Create(IdMessage.MessageParts, Anexo2);
        end;
      end;
      if (FormPrincipal.cd_empresa.Text = '10') then
      begin
        // Anexo da mensagem (opcional)
        if (FormPrincipal.rgTipo.ItemIndex = 0) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\10\Devolucao\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
        end;
        if (FormPrincipal.rgTipo.ItemIndex = 1) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\10\Garantia\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          Anexo2 := '\\192.168.0.2\Diversos\AUTORIZACOES\10\Garantia\' +
            'Ficha_Garantia.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
          TIdAttachmentFile.Create(IdMessage.MessageParts, Anexo2);
        end;
      end;
    except
      on E: Exception do
      begin
        raise Exception.Create
          ('Erro ao tenta anexao PDF Verifique se a pasta N:\Diversos esta mapeada: '
          + ' ' + E.Message);
        abort;
      end;
    end;

    // Envio da mensagem
    try
      IdSMTP.Send(IdMessage);
      MessageDlg('Mensagem enviada com sucesso.', mtInformation, [mbOK], 0);
      FormPrincipal.ETemp.Execute;
    except
      On E: Exception do
        MessageDlg('Erro ao enviar a mensagem: ' + E.Message, mtWarning,
          [mbOK], 0);
    end;
  finally
    // liberação dos objetos da memória
    FreeAndNil(IdMessage);
    FreeAndNil(IdSSLIOHandlerSocket);
    FreeAndNil(IdSMTP);
    FormSendEmail.Close;
  end;
end;

procedure TFormSendEmail.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FormPrincipal.bGerarPDF.Enabled :=true;
FormPrincipal.StatusBar1.Panels[2].Text := '';
  try
    FormPrincipal.FDQPDF.Close;
    DM.FDQ3.Close;
    DM.FDQ3.Active := FALSE;
    DM.FDQ3.SQl.Clear;
    DM.FDQ3.SQl.ADD('DELETE FROM espTempAutorizacao');
    DM.FDQ3.SQl.ADD
      ('WHERE (cd_empresa = :pcd_empresa) AND (nr_autoriz= :pnr_autoriz) AND (spid = :pspid)');
    DM.FDQ3.Params[0].AsString := FormPrincipal.cd_empresa.Text;
    DM.FDQ3.Params[1].AsString := FormPrincipal.nroAutorizacao.Text;
    DM.FDQ3.Params[2].ASInteger := DM.ConexaoUsuario;
    DM.FDQ3.ExecSQl;
  except
    on E: Exception do
    begin
      raise Exception.Create
        ('Erro ao tentar excluir os dados espTempAutorizacao: ' + ' ' +
        E.Message);
      abort;
    end;
  end;

  try
    DM.FDQ3.Close;
    DM.FDQ3.Active := FALSE;
    DM.FDQ3.SQl.Clear;
    DM.FDQ3.SQl.ADD('DELETE FROM espTempAutorizacaoItem');
    DM.FDQ3.SQl.ADD
      ('WHERE (cd_empresa = :pcd_empresa) AND (nr_autori= :pnr_autoriz) AND (spid = :pspid)');
    DM.FDQ3.Params[0].AsString := FormPrincipal.cd_empresa.Text;
    DM.FDQ3.Params[1].AsString := FormPrincipal.nroAutorizacao.Text;
    DM.FDQ3.Params[2].ASInteger := DM.ConexaoUsuario;
    DM.FDQ3.ExecSQl;
  except
    on E: Exception do
    begin
      raise Exception.Create
        ('Erro ao tentaar excluir os dados espTempAutorizacaoItem: ' + ' ' +
        E.Message);
      abort;
    end;
  end;

  try
    DM.FDQ3.Close;
    DM.FDQ3.Active := FALSE;
    DM.FDQ3.SQl.Clear;
    DM.FDQ3.SQl.ADD('DELETE FROM espTempPesqItem');
    DM.FDQ3.SQl.ADD
      ('WHERE (cd_empresa = :pcd_empresa) AND (nr_autoriz= :pnr_autoriz) AND (spid = :pspid)');
    DM.FDQ3.Params[0].AsString := FormPrincipal.cd_empresa.Text;
    DM.FDQ3.Params[1].AsString := FormPrincipal.nroAutorizacao.Text;
    DM.FDQ3.Params[2].ASInteger := DM.ConexaoUsuario;
    DM.FDQ3.ExecSQl;
  except
    on E: Exception do
    begin
      raise Exception.Create
        ('Erro ao tentaar excluir os dados espTempPesqItem: ' + ' ' +
        E.Message);
      abort;
    end;
  end;
end;

procedure TFormSendEmail.FormShow(Sender: TObject);
begin
self.Caption := 'Envio de Email [Spid: ' + IntToStr(DM.spid) +
      '/' + IntToStr(DM.ConexaoUsuario) + ' - Versão: ' + UFuncoes.versao_local + ']';
  edtRemetente.Text := FormPrincipal.FDQPDF.FieldByName('email_for').AsString;
  edtDestinatario.Text := Trim(FormPrincipal.FDQPDF.FieldByName('email_cli').AsString);
end;

procedure TFormSendEmail.SpeedButton1Click(Sender: TObject);
var
  // objetos necessários para o funcionamento
  IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL;
  IdSMTP: TIdSMTP;
  IdMessage: TIdMessage;
  CaminhoAnexo, Anexo2: string;
  x: string;
  i: integer;
const
  empresas: array [1 .. 4] of string = ('6', '7', '8', '14');
begin
  for i := 1 to 5 do
  begin
    if FormPrincipal.cd_empresa.Text = empresas[i] then
      x := empresas[i];
  end;
  // instanciação dos objetos
  IdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  IdSMTP := TIdSMTP.Create(Self);
  IdMessage := TIdMessage.Create(Self);
  // ShowMessage('teste 1');
  try
    // Configuração do SSL
    IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
    IdSSLIOHandlerSocket.SSLOptions.Mode := sslmClient;

    // Configuração do SMTP
    IdSMTP.IOHandler := IdSSLIOHandlerSocket;
    IdSMTP.AuthType := satDefault;
    IdSMTP.Port := 587;
    IdSMTP.Host := '****'; // aqui você coloca o mail.gmail.com ou outro que vai usar
    IdSMTP.Username := '****'; // aqui coloca o email
    IdSMTP.Password := '*******'; // aqui a senha
    // ShowMessage('teste 2');
    // Tentativa de conexão e autenticação
    try
      IdSMTP.Connect;
      IdSMTP.Authenticate;
    except
      on E: Exception do
      begin
        MessageDlg('Erro na conexão e/ou autenticação contatar o TI: ' +
          E.Message, mtWarning, [mbOK], 0);
        Exit;
      end;
    end;

    // Configuração da mensagem
    IdMessage.From.Address := edtRemetente.Text;
    IdMessage.From.Name := FormPrincipal.NomeEmpresa.Text;
    IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
    IdMessage.Recipients.EMailAddresses := edtDestinatario.Text;
    IdMessage.Subject := FormPrincipal.FDQPDF.FieldByName('msg_auto').AsString;
    IdMessage.Body.Text := 'Corpo do e-mail';

    try
      if (FormPrincipal.cd_empresa.Text = '1') then
      begin
        // Anexo da mensagem (opcional)
        if (FormPrincipal.rgTipo.ItemIndex = 0) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\1\Devolucao\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
        end;
        if (FormPrincipal.rgTipo.ItemIndex = 1) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\1\Garantia\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          Anexo2 := '\\192.168.0.2\Diversos\AUTORIZACOES\1\Garantia\' +
            'Ficha_Garantia.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
          TIdAttachmentFile.Create(IdMessage.MessageParts, Anexo2);
        end;
      end;
      if (FormPrincipal.cd_empresa.Text = x) then
      begin
        // Anexo da mensagem (opcional)
        if (FormPrincipal.rgTipo.ItemIndex = 0) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\' + x +
            '\Devolucao\' + 'Autorização - ' + FormPrincipal.nroAutorizacao.Text
            + ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
        end;
        if (FormPrincipal.rgTipo.ItemIndex = 1) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\' + x +
            '\Garantia\' + 'Autorização - ' + FormPrincipal.nroAutorizacao.Text
            + ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          Anexo2 := '\\192.168.0.2\Diversos\AUTORIZACOES\' + x + '\Garantia\' +
            'Ficha_Garantia.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
          TIdAttachmentFile.Create(IdMessage.MessageParts, Anexo2);
        end;
      end;
      if (FormPrincipal.cd_empresa.Text = '10') then
      begin
        // Anexo da mensagem (opcional)
        if (FormPrincipal.rgTipo.ItemIndex = 0) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\10\Devolucao\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
        end;
        if (FormPrincipal.rgTipo.ItemIndex = 1) then
        begin
          CaminhoAnexo := '\\192.168.0.2\Diversos\AUTORIZACOES\10\Garantia\' +
            'Autorização - ' + FormPrincipal.nroAutorizacao.Text +
            ' - Cliente - ' + FormPrincipal.cd_cliente.Text + '.pdf';
          Anexo2 := '\\192.168.0.2\Diversos\AUTORIZACOES\10\Garantia\' +
            'Ficha_Garantia.pdf';
          if FileExists(CaminhoAnexo) then
            TIdAttachmentFile.Create(IdMessage.MessageParts, CaminhoAnexo);
          TIdAttachmentFile.Create(IdMessage.MessageParts, Anexo2);
        end;
      end;
    except
      on E: Exception do
      begin
        raise Exception.Create
          ('Erro ao tenta anexao PDF Verifique se a pasta N:\Diversos esta mapeada: '
          + ' ' + E.Message);
        abort;
      end;
    end;

    // Envio da mensagem
    try
      IdSMTP.Send(IdMessage);
      MessageDlg('Mensagem enviada com sucesso.', mtInformation, [mbOK], 0);
    except
      On E: Exception do
        MessageDlg('Erro ao enviar a mensagem: ' + E.Message, mtWarning,
          [mbOK], 0);
    end;
  finally
    // liberação dos objetos da memória
    FreeAndNil(IdMessage);
    FreeAndNil(IdSSLIOHandlerSocket);
    FreeAndNil(IdSMTP);
    FormSendEmail.Close;
  end;

end;

end.
