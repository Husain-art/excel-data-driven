package hcyclewala;

import java.io.IOException;
import java.util.ArrayList;

public class ExcelData {
public static void main(String[] args) throws IOException {
    dataDriven dd = new dataDriven();
    ArrayList<String> a = dd.getData("Purchase");

    System.out.println(a.get(2));

}
}
