object TableFrame: TTableFrame
  Left = 0
  Top = 0
  Width = 674
  Height = 526
  ParentShowHint = False
  ShowHint = False
  TabOrder = 0
  object sgTable: TStringGrid
    Left = 0
    Top = 31
    Width = 674
    Height = 495
    Align = alClient
    DefaultColWidth = 10
    DefaultRowHeight = 18
    DrawingStyle = gdsGradient
    RowCount = 3
    GradientEndColor = 15790320
    GradientStartColor = 15790320
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing, goEditing, goThumbTracking]
    PopupMenu = pmGrid
    TabOrder = 0
    OnContextPopup = sgTableContextPopup
    OnDrawCell = sgTableDrawCell
    OnGetEditText = sgTableGetEditText
    OnKeyDown = sgTableKeyDown
    OnMouseWheelDown = sgTableMouseWheelDown
    OnMouseWheelUp = sgTableMouseWheelUp
    OnSelectCell = sgTableSelectCell
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 674
    Height = 31
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object cbFixColumns: TCheckBox
      Left = 8
      Top = 6
      Width = 177
      Height = 17
      Caption = 'Lock first              Columns'
      TabOrder = 0
      OnClick = cbFixColumnsClick
    end
    object seFixedColumns: TSpinEdit
      Left = 72
      Top = 4
      Width = 33
      Height = 22
      Enabled = False
      MaxValue = 9
      MinValue = 1
      TabOrder = 1
      Value = 1
      OnChange = cbFixColumnsClick
    end
  end
  object pmGrid: TPopupMenu
    Left = 392
    Top = 160
    object miGridAddNew: TMenuItem
      Caption = 'Append'
      OnClick = miGridAddNewClick
    end
    object miGridAddCopy: TMenuItem
      Caption = 'Append copy'
      OnClick = miGridAddCopyClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object miGridInsertNew: TMenuItem
      Caption = 'Insert'
      OnClick = miGridInsertNewClick
    end
    object miGridInsertCopy: TMenuItem
      Caption = 'Insert copy'
      OnClick = miGridInsertCopyClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object miGridDeleteRow: TMenuItem
      Caption = 'Delete'
      OnClick = miGridDeleteRowClick
    end
  end
end
