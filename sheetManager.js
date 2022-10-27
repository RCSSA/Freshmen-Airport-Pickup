class sheetManager{
    constructor(sheet, colMap){
      this.sheet = sheet
      this.colMap = colMap
    }
  
    // Contract:
    // Input params:
    // 1. col is the name of the column
    // Output:
    // A list of values
    getValuesByCol(col){
      // the first row of every sheet is a header
      // so it's ruled out under all value involved cases.
      return this.sheet.getRange(2, this.colMap[col],            // top left coordinate of the recetangular
                                 this.sheet.getLastRow() - 1,    // height of the recetangular
                                 1)                              // width of the recetangular
                                 .getValues().flat()
    }
  
    // Contract:
    // Input params:
    // 1. col is the name of the column
    // Output:
    // A Range
    getByCol(col){
      // the first row of every sheet is a header
      // so it's ruled out under all value involved cases.
      return this.sheet.getRange(2, this.colMap[col],            // top left coordinate of the recetangular
                                 this.sheet.getLastRow() - 1,    // height of the recetangular
                                 1)                              // width of the recetangular
    }
  
    // Contract:
    // Input params:
    // 1. pos is the row index
    // Output:
    // A list of values
    getValuesByRow(pos){
      return this.sheet.getRange(pos, 1,                          // top left coordinate of the recetangular
                                 1,                               // height of the recetangular
                                 this.sheet.getLastColumn())      // width of the recetangular
                                 .getValues().flat()
    }
  
    // Contract:
    // Input params:
    // 1. pos is the row index
    // Output:
    // A Range
    getByRow(pos){
      return this.sheet.getRange(pos, 1,                          // top left coordinate of the recetangular
                                 1,                               // height of the recetangular
                                 this.sheet.getLastColumn())      // width of the recetangular
    }
  
    // Contract:
    // Input params:
    // 1. pos is the row index of the target place
    // 2. obj is an object or a map that contains all required fields of the sheet
    insert(pos, obj){
      // insert an empty row
      this.sheet.insertRowAfter(pos)
  
      this.updateRow(pos, obj)
    }
  
    // Contract:
    // Input Params:
    // 1. pos is the row index of the target
    // 2. obj is an object or a map that contains all required fields of the sheet
    updateRow(pos, obj){
      let valueList = []
      let cols = this.colMap.keys()
      while(!cols.done){
        valueList.push(obj[cols.value])
        cols = cols.next()
      }
      
      this.getByRow(pos).setValues([valueList])
    }
    
    // Contract:
    // Input Params:
    // 1. pos is the row index of the target
    // 2. col is the name of the column of the target
    // 3. value
    updateRowProperty(pos, col, value){
      this.sheet.getRange(pos, this.colMap[col], 1, 1).setValue(value)
    }
  
    // Contract:
    // Input Params:
    // 1. col is the name of the column
    // 2. values is a list of values
    updateCol(col, values){
      // `setValues` require input to be list of lists
      let valueList = []
      for(let i = 0; i < values.length; i++){
        valueList.push([values[i]])
      }
  
      this.getByCol(col).setValues(valueList)
    }
  
    // Contract:
    // Input Params:
    // 1. col is the name of the column
    // 2. searchString supports:
    // * string
    // * date object
    searchInCol(col, searchString){
      if (searchString instanceof Date){
        let _tmp = [(searchString.getMonth() + 1).toString(),   // Date.getMonth()'s retrun value is 0 based, so add one here
                     searchString.getDate().toString(), 
                     searchString.getFullYear().toString()].join('/')
        searchString = _tmp
      }
  
      
      let textFinder = this.getByCol(col).createTextFinder(searchString)
      return textFinder.findNext()
    }
  }