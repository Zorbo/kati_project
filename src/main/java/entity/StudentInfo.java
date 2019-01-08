package entity;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class StudentInfo {

    private String time;

    private String name;

    private String classNumber;

    private String reason;

}
