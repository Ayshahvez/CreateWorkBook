import java.io.IOException;

public class CreateWorkBook
{
    public static void main(String[] args) {
        try {
            new MainWindow();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}