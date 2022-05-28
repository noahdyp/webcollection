package per.dyp.webcollection;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jsoup.Connection;
import org.jsoup.Connection.Response;
import org.jsoup.Jsoup;

import per.dyp.webcollection.QiuTanExporter.Match;

public class QiuTanCollector implements Collector {

    private static final String BFDATA_URL = "http://bf.titan007.com/vbsxml/bfdata.js";

    // http://1x2.nowscore.com/1292139.js
    private static final String MATCH_URL_FORMAT = "http://1x2.nowscore.com/%d.js";

    private static final Pattern SCORE_ID_PATTERN = Pattern
            .compile("A\\[\\d+\\]=\"(\\d+)\\^([^\\^]*\\^){12}([0-4])\\^");

    private static final Pattern MATCH_INFO_PATTERN = Pattern.compile("(\\S+)=(\"?)((?!;\").*)\\2;");

    public Object collect() {

        List<Match> result = new ArrayList<Match>();
        try {
            Connection con = Jsoup.connect(BFDATA_URL);
            Response res = con.ignoreContentType(true).execute();
            if (res != null && res.bodyAsBytes() != null) {
                String bfdata = new String(res.bodyAsBytes(), "gbk");
                Matcher scoreMatcher = SCORE_ID_PATTERN.matcher(bfdata);
                List<Long> scoreIds = new ArrayList<Long>();
                while (scoreMatcher.find()) {
                    scoreIds.add(Long.valueOf(scoreMatcher.group(1)));
                }
                System.out.println("比赛场次 : " + scoreIds.size());
                for (Long scoreId : scoreIds) {
                    try {
                        String matchUrl = String.format(MATCH_URL_FORMAT, scoreId);
                        Connection matchCon = Jsoup.connect(matchUrl);
                        Response matchRes = matchCon.ignoreContentType(true).execute();
                        if (matchRes != null && matchRes.bodyAsBytes() != null) {
                            String matchInfo = new String(matchRes.bodyAsBytes(), "UTF-8");
                            Matcher matcher = MATCH_INFO_PATTERN.matcher(matchInfo);
                            Map<String, String> params = new HashMap<String, String>();
                            while (matcher.find()) {
                                params.put(matcher.group(1), matcher.group(3));
                            }
                            Match match = new Match(scoreId, params);
                            result.add(match);
                        }
                    } catch (Exception e) {
                        // System.out.println(e.getMessage());
                    }
                }
                System.out.println("开盘的场次 : " + result.size());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

}
