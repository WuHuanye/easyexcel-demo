package easyexceldemo.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import easyexceldemo.dto.BizMergeStrategy;
import easyexceldemo.dto.RowRangeDto;
import easyexceldemo.dto.TitleSheetWriteHandler;
import easyexceldemo.entity.User;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * <p>
 * 该类的描述
 * </p>
 *
 * @author
 * @Modified By:
 * @since 2020/11/23 11:31
 */
@RestController
@RequestMapping("/user")
public class EasyExcelController {

    @GetMapping("/excel")
    public void excel(HttpServletResponse response) throws IOException {
        Map<String, List<RowRangeDto>> strategyMap = BizMergeStrategy.addAnnualMerStrategy(data());
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String filename = URLEncoder.encode("用户表测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename=" + filename + ".xlsx");

            EasyExcel.write(response.getOutputStream(), User.class)
                    .excelType(ExcelTypeEnum.XLSX).head(User.class)
                    .registerWriteHandler(new TitleSheetWriteHandler("我是一个小标题", 2)) // 标题及样式，lastCol为标题第0列到底lastCol列的宽度
                    //设置默认样式及写入头信息开始的行数
                    .relativeHeadRowIndex(1)
                    // 注册合并策略
                    .registerWriteHandler(new BizMergeStrategy(strategyMap))
                    // 设置样式
                    .registerWriteHandler(BizMergeStrategy.CellStyleStrategy())
                    .sheet("测试")
                    .doWrite(data());
        } catch (Exception e) {
            e.printStackTrace();
            response.reset();
            response.setCharacterEncoding("utf-8");
            response.setContentType("application/json");
            response.getWriter().println("打印失败");
        }

    }

    private List<User> data() {
        List<User> list = new ArrayList<>();
        User user = new User("1", "张三", "总裁");
        User user1 = new User("2", "李四", "总经理");
        User user2 = new User("3", "李四", "技术员");
        User user3 = new User("4", "王五", "技术员");

        list.add(user);
        list.add(user1);
        list.add(user2);
        list.add(user3);

        return list;
    }
}
