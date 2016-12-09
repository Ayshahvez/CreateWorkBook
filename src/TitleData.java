/**
 * Created by Ayshahvez on 11/11/2016.
 */
public class TitleData {
    public TitleData(String sheetTitle) {
        super();
        SheetTitle = sheetTitle;
    }

    public TitleData(){
        super();
    }

    public String getSheetTitle() {

        return SheetTitle;
    }

    public void setSheetTitle(String sheetTitle) {
        SheetTitle = sheetTitle;
    }

    private String SheetTitle;

}
