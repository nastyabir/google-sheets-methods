def hide_grid(s, worksheet_id, hidegrid=True):
    request = {
    "requests": [
      {
        "updateSheetProperties": {
          "fields": "gridProperties(hideGridlines)",    
          "properties": {
            "sheetId":f"{worksheet_id}",
            "gridProperties": {
              "hideGridlines": hidegrid
            }
          }
        }
      }
    ],
    "includeSpreadsheetInResponse": False,
    "responseIncludeGridData": False,
  }
    s.batch_update(request)
    

def merge_cells(s, worksheet_id, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex, mergeType="MERGE_ALL"):
    request = {
  "requests": [
    {
      "mergeCells": {
        "range": {
          "sheetId": f"{worksheet_id}",
          "startRowIndex": startRowIndex,
          "endRowIndex": endRowIndex,
          "startColumnIndex": startColumnIndex,
          "endColumnIndex": endColumnIndex
        },
        "mergeType": mergeType
      }
    },
    
  ]
}
    s.batch_update(request)
    
    
def wrap_cell(s, worksheet_id, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex, wrapStrategy="WRAP"):
    request = {
  "requests": [
    {
      "updateCells": {
        "range": {
          "sheetId": f"{worksheet_id}",
          "startRowIndex": startRowIndex,
          "endRowIndex": endRowIndex,
          "startColumnIndex": startColumnIndex,
          "endColumnIndex": endColumnIndex
        },
        "rows": [
          {
            "values": [
              {
                "userEnteredFormat": {
                  "wrapStrategy": wrapStrategy
                }
              }
            ]
          }
        ],
        "fields": "userEnteredFormat.wrapStrategy"
      }
    }
  ]
}
    s.batch_update(request)
    
    
def resize_cell(s, worksheet_id, dimension, startIndex, endIndex, pixelSize):
    request = {
  "requests": [{
    "updateDimensionProperties": {
    "range": {
      "sheetId": f"{worksheet_id}",
      "dimension": dimension,
      "startIndex": startIndex,
      "endIndex": endIndex
    },
    "properties": {
      "pixelSize": pixelSize
    },
    "fields": "pixelSize"
    }
}
  ]
}
    s.batch_update(request)
    
    
def cell_value(worksheet, row, col, value):
    new_cell = worksheet.cell(row=row, col=col)
    new_cell.value = value
    worksheet.update_cells([new_cell]) 
