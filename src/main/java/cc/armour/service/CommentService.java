package cc.armour.service;

import cc.armour.entity.Comment;
import cc.armour.util.EncryptUtils;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONPath;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import java.math.BigInteger;
import java.net.URL;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CommentService {

    private static final String DOMAIN = "http://music.163.com";

    private static final String NET_EASE_COMMENT_API_URL = "https://music.163.com/weapi/comment/resource/comments/get?csrf_token=";

    //通过歌曲ID获取评论API，网易对其进行了加密
    public static List<Comment> parseCommentMessage(String songId) throws Exception {
        String songUrl = DOMAIN + "/song?id=" + songId;
        long pageSize = 0;
        String cursor = "-1";
        List<Comment> comments = new ArrayList<>();
        for (long commentIndex = 0; commentIndex <= pageSize;commentIndex++ ) {
            Thread.sleep((long) (Math.random() * 500));
            String secKey = new BigInteger(100, new SecureRandom()).toString(32).substring(0, 16);
            String encText = EncryptUtils.aesEncrypt(EncryptUtils.aesEncrypt("{\"csrf_token\": \"\", \"cursor\":\"" + cursor +"\", \"offset\": \""+ 0 + "\", \"orderType\": \"1\", \"pageNo\":\""+ (commentIndex+1) +"\",\"pageSize\": \"50\", \"rid\": \"R_SO_4_" + songId +"\", \"threadId\": \"R_SO_4_"+ songId+"\"}", "0CoJUm6Qyw8W8jud"), secKey);
            String encSecKey = EncryptUtils.rsaEncrypt(secKey);
            Map<String, String> paramMap = new HashMap<>();
            paramMap.put("params", encText);
            paramMap.put("encSecKey", encSecKey);
            Connection.Response response = Jsoup
                    .connect(NET_EASE_COMMENT_API_URL)
                    .ignoreContentType(true)
                    .header("User-Agent",
                            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36")
                    .header("Accept-Language", "zh-CN,zh;q=0.8,en;q=0.6").header("Connection", "keep-alive")
                    .header("Referer", "https://music.163.com/song?id="+songId)
                    .header("Origin", "http://music.163.com").header("Host", "music.163.com")
                    .header("Content-Type", "application/x-www-form-urlencoded")
                    .method(Connection.Method.POST)
                    .data(paramMap).execute();
            Object res = JSON.parse(response.body());
            if(response.body().isEmpty()) {
                commentIndex--;
                continue;
            }
            int code = (int) JSONPath.eval(res, "$.code");
            if (200 != code) {
                Thread.sleep((long) (Math.random() * 5000));
                System.out.println("休息一会儿"+commentIndex);
                commentIndex--;
                continue;
            }
            int commentCount = (int) JSONPath.eval(res, "$.data.totalCount");
            int curruntCommentCount = (int) JSONPath.eval(res, "$.data.comments.size()");
            cursor = (String) JSONPath.eval(res, "$.data.cursor");
            pageSize = commentCount / 50;
            if(curruntCommentCount==0) {commentIndex--;continue;}
            for (int i = 0; i < curruntCommentCount; i++) {
                String nickname = JSONPath.eval(res, "$.data.comments[" + i + "].user.nickname").toString();
                String time = EncryptUtils.stampToDate((long) JSONPath.eval(res, "$.data.comments[" + i + "].time"));
                String content = JSONPath.eval(res, "$.data.comments[" + i + "].content").toString();
                String likedCount = JSONPath.eval(res, "$.data.comments[" + i + "].likedCount").toString();
                comments.add(new Comment(content, nickname, time, Integer.parseInt(likedCount)));
            }
        }

        return comments;
    }
}
