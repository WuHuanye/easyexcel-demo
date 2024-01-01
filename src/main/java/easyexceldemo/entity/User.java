package easyexceldemo.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;

/**
 * <p>
 * 该类的描述
 * </p>
 *
 * @author
 * @Modified By:
 * @since 2020/11/23 11:28
 */
public class User {
    @ColumnWidth(10)
    @ExcelProperty(value = {"id"}, index = 0)
    private String id;

    @ColumnWidth(20)
    @ExcelProperty(value = {"名字"}, index = 1)
    private String name;

    @ColumnWidth(20)
    @ExcelProperty(value = {"职位"}, index = 2)
    private String postion; //职位

    public User(String id,String name,String postion){
        this.id = id;
        this.name = name;
        this.postion = postion;
    }
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPostion() {
        return postion;
    }

    public void setPostion(String postion) {
        this.postion = postion;
    }
}
