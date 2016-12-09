import java.util.ArrayList;
import java.util.List;

/**
 * Created by Ayshahvez on 11/11/2016.
 */
public class TitleModel {
    public List<TitleData> findAll() {
        try{
            ArrayList<TitleData> result = new ArrayList<TitleData>();
            result.add(new TitleData("ABC PENSION FUND"));
            result.add(new TitleData("ACTUARIAL FUNDING VALUATION AS AT 2008 MARCH 31"));
            result.add(new TitleData("ACCUMULATION OF ACTIVE MEMBERS' ACCOUNT BALANCES"));

            return result;
        }
        catch(Exception e){
            return null;
        }
    }
}
