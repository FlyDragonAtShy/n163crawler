package cc.armour.entity;

import cc.armour.annotation.ExcelColumn;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.experimental.Accessors;

@Data
@AllArgsConstructor
@Accessors(chain = true)
public class Comment {

    @ExcelColumn(value = "评论内容", col = 2)
    private String content;

    @ExcelColumn(value = "用户昵称", col = 1)
    private String nickname;

    @ExcelColumn(value = "评论时间", col = 3)
    private String time;

    @ExcelColumn(value = "点赞数", col = 4)
    private int likedCount;

}
