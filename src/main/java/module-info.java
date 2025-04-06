module com.example.ishomework {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;


    opens com.example.ishomework to javafx.fxml;
    exports com.example.ishomework;
}