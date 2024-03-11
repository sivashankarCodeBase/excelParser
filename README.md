# excelParser

## A minimalistic excel parser

### Inputs:
- File name
- Sheet name

### Outputs:
- map[string][]string
- error

### Example:

```go
data, err := excel.Read("test1.xlsx", "sheet1")
if err != nil {
    fmt.Println(err)
}
fmt.Println(data)
