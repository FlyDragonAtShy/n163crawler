package cc.armour.main;

import cc.armour.entity.Comment;
import cc.armour.service.CommentService;
import cc.armour.util.ExcelUtils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

public class RunProgram {
    public static void main(String[] args) throws Exception {
        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入歌曲名-歌曲id");
        String in = scanner.next();
        String songName = in.split("-")[0];
        String songId = in.split("-")[1];
        List<Comment> commentList = CommentService.parseCommentMessage(songId);
        HashMap map = new HashMap();
        map.put(songName, commentList);
        ExcelUtils.writeExcel("/n163crawler/excel/"+songName+".xls", map, Comment.class);
    }
}
