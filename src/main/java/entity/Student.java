package entity;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class Student {

    private String time;

    private String name;

    private String classNumber;

    private String reason;

    private String key;
}
