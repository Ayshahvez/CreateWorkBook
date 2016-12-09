import java.util.ArrayList;
import java.util.List;

/**
 * Created by Ayshahvez on 11/10/2016.
 */
public class MemberModel {

    public List<MemberInfo> findAll() {
       try{
           ArrayList<MemberInfo> result = new ArrayList<MemberInfo>();
         //  result.add(new MemberInfo("1242572","Iraida ","Sheard","f"," ","01/23/1968","02/12/2000","03/01/2000"));
         //  result.add(new MemberInfo("p2","Baily","Chung","05/23/1988"));
         //  result.add(new MemberInfo("p3","Rob","Parker","08/23/1965"));

           return result;
       }
       catch(Exception e){
           return null;
       }
    }
}
