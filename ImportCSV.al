xmlport 50100 "Import Items XMLport"
{
    Caption = 'Import Domains';
    Format = VariableText;
    Direction = Import;
    TextEncoding = UTF8;
    UseRequestPage = false;
    FileName = 'ItemsViaXMLport.csv';
    TableSeparator = '<NewLine>';

    schema
    {
        textelement(Root)
        {
             tableelement(Domain; Domain)
            {
                RequestFilterFields = "Domain /Subscription";
                fieldelement(Domains; Domain."Domain /Subscription")
                { }
                fieldelement(NichtInAutoDNS; Domain."Not in AutoDNS")
                { }
                fieldelement(Debitorennr; Domain."Debtor no.")
                { }
                fieldelement(Debitorenname; Domain."Debitor name")
                { }
                fieldelement(NächstAbrechungsZeitraum; Domain."Next billing date")
                { }
                fieldelement(Monatspreis; Domain."Monthly price")
                { }
                fieldelement(Abrechnungszeitraum; Domain."Billing period")
                { }
                fieldelement(Registrierungsdatum; Domain."Registration date")
                { }
                fieldelement(Vertragsende; Domain."End of contract")
                { }
                fieldelement(AliasDomain; Domain."Alias-domain")
                { }
                fieldelement(Status; Domain.Status)
                { }
                fieldelement(Kündigungsdatum; Domain."Termination date")
                { }
                fieldelement(AngelegtVon; Domain."Created By")
                { }
                fieldelement(AbgerechnetBisZum; Domain."Settled until")
                { }
                fieldelement(Kommentar; Domain.comment)
                { }
                fieldelement(Vereinbarung; Domain.Argeement)
                { }
                fieldelement(Art; Domain.Type)
                { }
                fieldelement(VonPZerstellt; Domain."Created by PZ")
                { }

                trigger OnAfterInitRecord()
                begin
                    if IsFirstline then begin
                        IsFirstline := false;
                        currXMLport.Skip();
                    end;
                end;

                trigger OnAfterGetRecord()
                var
                    domain: Record Domain;

                begin
                    domain.Reset();
                    domain.SetRange("Domain /Subscription", Domain."Domain /Subscription");
                    if domain.FindFirst() then domain.Delete(true);
                end;
            }
        }
    }

    trigger OnPreXmlPort()
    begin
        IsFirstline := true;
    end;

    var
        IsFirstline: Boolean;
}

pageextension 50100 DomainExt extends "Domain List"
{
    actions
    {
        addafter(History)
        {
            action(ImportItemsviaXMLport)
            {
                Caption = 'Import Domain via XMLport';
                Promoted = true;
                PromotedCategory = Process;
                Image = Import;
                ApplicationArea = All;

                trigger OnAction()
                begin
                    Xmlport.Run(Xmlport::"Import Items XmlPort", false, true);
                end;
            }

            action(ImportItemsviaCSVBuffer)
            {
                Caption = 'Import Items via CSV Buffer';
                Promoted = true;
                PromotedCategory = Process;
                Image = Import;
                ApplicationArea = All;

                trigger OnAction()
                var
                    InS: InStream;
                    FileName: Text[100];
                    UploadMsg: Label 'Please choose the CSV file';
                    Item: Record Item;
                    LineNo: Integer;
                begin
                    CSVBuffer.Reset();
                    CSVBuffer.DeleteAll();
                    if UploadIntoStream(UploadMsg, '', '', FileName, InS) then begin
                        CSVBuffer.LoadDataFromStream(InS, ',');
                        for LineNo := 2 to CSVBuffer.GetNumberOfLines() do begin
                            Item.Init();
                            Item.Validate("No.", GetValueAtCell(LineNo, 1));
                            Item.Insert(true);
                            Item.Validate(Description, GetValueAtCell(LineNo, 2));
                            case GetValueAtCell(LineNo, 3) of
                                'Inventory':
                                    Item.Validate(Type, Item.Type::"Inventory");
                                'Service':
                                    Item.Validate(Type, Item.Type::"Service");
                                'Non-Inventory':
                                    Item.Validate(Type, Item.Type::"Non-Inventory");
                            end;
                            Item.Validate("Base Unit of Measure", GetValueAtCell(LineNo, 5));
                            Evaluate(Item."Unit Cost", GetValueAtCell(LineNo, 7));
                            Evaluate(Item."Unit Price", GetValueAtCell(LineNo, 8));
                            Item.Validate("Vendor No.", GetValueAtCell(LineNo, 9));
                            Item.Modify(true);
                        end;
                    end;
                end;
            }
        }
    }

    var
        CSVBuffer: Record "CSV Buffer" temporary;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin
        if CSVBuffer.Get(RowNo, ColNo) then
            exit(CSVBuffer.Value)
        else
            exit('');
    end;
}
