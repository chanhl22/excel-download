# excel-download
* 엑셀로 만들 DTO 객체에 어노테이션을 붙여서 바로 적용할 수 있습니다. 필요에 따라 cell style을 커스텀한 클래스를 구현하면 됩니다.
* 예시 참고
```java
@NoArgsConstructor
@AllArgsConstructor
@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class))
@DefaultBodyStyle(style = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class))
public class SampleExcelDto {

    @ExcelColumn(
            headerName = "번호",
            headerStyle = @ExcelColumnStyle(excelCellStyleClass = BlueHeaderStyle.class),
            bodyStyle = @ExcelColumnStyle(excelCellStyleClass = AlignCenterAndBordersThinBodyStyle.class)
    )
    private long id;

    @ExcelColumn(headerName = "이름")
    private String name;

    @ExcelColumn(headerName = "제목")
    private String title;
}
```
1. @DefaultHeaderStyle
    * 엑셀 전체 헤더에 적용되는 어노테이션입니다.
    * style 속성을 지정해줘야 합니다.
2. @ExcelColumnStyle
    * 엑셀에 style 속성을 적용하는 어노테이션입니다.
    * 직접 class 를 만들어서 excelCellStyleClass 속성에 넣어주면 됩니다.
    * NoExcelCellStyle.class 는 아무것도 적용하지 않는 커스텀 스타일 클래스입니다.
       ```java
       public class NoExcelCellStyle implements ExcelCellStyle {
 
           @Override
           public void apply(CellStyle cellStyle) {
                // Do nothing
           }
       }
       ```  
3. @DefaultBodyStyle
    * 엑셀 전체 바디에 적용되는 어노테이션입니다.
    * style 속성을 지정해줘야 합니다.
4. @ExcelColumn
    * 각 칼럼에 적용되는 어노테이션입니다.
    * headerName, headerStyle, bodyStyle 속성이 있습니다.
        1. headerName: 칼럼의 헤더 이름, default = ""
        2. headerStyle: 칼럼의 헤더 스타일, default = NoExcelCellStyle.class
        3. bodyStyle: 칼럼의 바디 스타일, default = NoExcelCellStyle.class
    * BlueHeaderStyle.class 는 중앙 정렬, 파란 배경, 얇은 테두리를 적용하는 커스텀 스타일 클래스입니다.
      ```java
      public class BlueHeaderStyle extends CustomExcelCellStyle {
    
          @Override
          public void configure(ExcelCellStyleConfigurer configurer) {
              configurer
                    .excelAlign(DefaultExcelAlign.CENTER_CENTER)
                    .foregroundColor(223, 235, 246)
                    .excelBorders(DefaultExcelBorders.newInstance(ExcelBorderStyle.THIN));
          }
      }
      ```
    * AlignCenterAndBordersThinBodyStyle.class 는 중앙 정렬, 얇은 테두리를 적용하는 커스텀 스타일 클래스입니다.
      ```java
      public class AlignCenterAndBordersThinBodyStyle extends CustomExcelCellStyle {
 
          @Override
          public void configure(ExcelCellStyleConfigurer configurer) {
              configurer
                    .excelAlign(DefaultExcelAlign.CENTER_CENTER)
                    .foregroundColor(255, 255, 255)
                    .excelBorders(DefaultExcelBorders.newInstance(ExcelBorderStyle.THIN));
         }
      }
      ```
5. 샘플 코드 이미지  
   ![엑셀 이미지](/README/img_1.png)
    * DTO의 각 필드의 헤더 이름이 설정한 번호, 이름, 제목으로 변경됐습니다.
    * DTO의 id 필드의 헤더 스타일이 중앙 정렬, 파란 배경, 얇은 테두리로 변경됐습니다.
    * DTO의 id 필드의 바디 스타일이 중앙 정렬, 얇은 테두리로 변경됐습니다.