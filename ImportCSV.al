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
            tableelement(Item; Item)
            {
                XmlName = 'Domains';
                RequestFilterFields = "No.";
                fieldelement(No; Item."No.")
                {
                }
                fieldelement(Description; Item.Description)
                {
                }
                fieldelement(Type; Item.Type)
                {
                }
                fieldelement(Inventory; Item.Inventory)
                {
                }
                fieldelement(BaseUnitofMeasure; Item."Base Unit of Measure")
                {
                }
                fieldelement(CostisAdjusted; Item."Cost is Adjusted")
                {
                }
                fieldelement(UnitCost; Item."Unit Cost")
                {
                }
                fieldelement(UnitPrice; Item."Unit Price")
                {
                }
                fieldelement(VendorNo; Item."Vendor No.")
                {
                }

                trigger OnAfterInitRecord()
                begin
                    if IsFirstline then begin
                        IsFirstline := false;
                        currXMLport.Skip();
                    end;
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

pageextension 50100 ItemExt extends "Item List"
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
