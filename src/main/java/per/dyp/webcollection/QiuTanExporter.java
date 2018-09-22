package per.dyp.webcollection;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import per.dyp.webcollection.common.DateUtils;

public class QiuTanExporter implements Exporter {

    private static final String FILE_NAME_FORMAT = "QiuTan(%s).xlsx";
    private static final int perCellMatchCount = 5;
    private final Map<String, XSSFCellStyle> styleMap = new HashMap<String, XSSFCellStyle>();

    public void export(Object object, String path) {
        try {
            List<Match> matchs;
            if (object instanceof List) {
                matchs = (List<Match>) object;
            } else {
                return;
            }
            XSSFWorkbook xssf = getXSSFWorkbook();
            XSSFSheet sheet = xssf.createSheet("sheet");
            sheet.setDefaultRowHeight((short) (14 * 20));
            sheet.setDefaultColumnWidth(8);

            for (int i = 0; i < perCellMatchCount * 7; i++) {
                sheet.createRow(i);
            }
            int matchCount = matchs.size();
            int loopTimes = matchCount / perCellMatchCount;
            for (int times = 0; times < loopTimes; times++) {
                int cellIndex = times * 13;
                for (int index = 0; index < perCellMatchCount; index++) {
                    int rowIndex = index * 7;
                    Match match = matchs.get(perCellMatchCount * times + index);
                    buildXlsx(sheet, cellIndex, rowIndex, match);
                }
            }
            int cellIndex = loopTimes * 13;
            int extraTimes = matchCount % perCellMatchCount;
            for (int index = 0; index < extraTimes; index++) {
                int rowIndex = index * 7;
                Match match = matchs.get(perCellMatchCount * loopTimes + index);
                buildXlsx(sheet, cellIndex, rowIndex, match);
            }
            String fileName = String.format(FILE_NAME_FORMAT, DateUtils.getCurrentDateTime(DateUtils.yyyyMMddHHmmss));
            File file = new File(path + fileName);
            if (!file.exists()) {
                file.createNewFile();
            }
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            xssf.write(os);
            byte[] xls = os.toByteArray();

            OutputStream out = new FileOutputStream(file);
            out.write(xls);
            out.close();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private XSSFWorkbook getXSSFWorkbook() {
        XSSFWorkbook xssf = new XSSFWorkbook();

        XSSFCellStyle common = xssf.createCellStyle();
        common.setAlignment(HorizontalAlignment.CENTER);
        common.setVerticalAlignment(VerticalAlignment.CENTER);
        // common.setWrapText(true);

        XSSFFont name = xssf.createFont();
        name.setFontName("Tahoma");
        name.setFontHeight(12);
        name.setBold(true);
        XSSFCellStyle nameStyle = ((XSSFCellStyle) (common.clone()));
        nameStyle.setFont(name);
        styleMap.put("name", nameStyle);

        XSSFFont kaiLi = xssf.createFont();
        kaiLi.setFontName("Tahoma");
        kaiLi.setFontHeight(12);
        XSSFCellStyle kaiLiStyle = ((XSSFCellStyle) (common.clone()));
        kaiLiStyle.setFont(kaiLi);
        styleMap.put("kaiLi", kaiLiStyle);

        XSSFFont peiLvMoreThanChuShi = xssf.createFont();
        peiLvMoreThanChuShi.setFontName("Tahoma");
        peiLvMoreThanChuShi.setFontHeight(12);
        peiLvMoreThanChuShi.setBold(true);
        peiLvMoreThanChuShi.setColor(IndexedColors.RED.index);
        XSSFCellStyle peiLvMoreThanChuShiStyle = ((XSSFCellStyle) (common.clone()));
        peiLvMoreThanChuShiStyle.setFont(peiLvMoreThanChuShi);
        styleMap.put("peiLvMoreThanChuShi", peiLvMoreThanChuShiStyle);

        XSSFFont peiLvLessThanChuShi = xssf.createFont();
        peiLvLessThanChuShi.setFontName("Tahoma");
        peiLvLessThanChuShi.setFontHeight(12);
        peiLvLessThanChuShi.setBold(true);
        peiLvLessThanChuShi.setColor(IndexedColors.BRIGHT_GREEN.index);
        XSSFCellStyle peiLvLessThanChuShiStyle = ((XSSFCellStyle) (common.clone()));
        peiLvLessThanChuShiStyle.setFont(peiLvLessThanChuShi);
        styleMap.put("peiLvLessThanChuShi", peiLvLessThanChuShiStyle);

        XSSFFont peiLvSubtraNearZero = xssf.createFont();
        peiLvSubtraNearZero.setFontName("Tahoma");
        peiLvSubtraNearZero.setFontHeight(12);
        peiLvSubtraNearZero.setBold(true);
        peiLvSubtraNearZero.setColor(IndexedColors.RED.index);
        XSSFCellStyle peiLvSubtraNearZeroStyle = ((XSSFCellStyle) (common.clone()));
        peiLvSubtraNearZeroStyle.setFont(peiLvSubtraNearZero);
        styleMap.put("peiLvSubtraNearZero", peiLvSubtraNearZeroStyle);

        return xssf;
    }

    private void buildXlsx(XSSFSheet sheet, int cellIndex, int rowIndex, Match match) {

        XSSFRow matchTeam = sheet.getRow(rowIndex);
        CellRangeAddress craMatchTeam = new CellRangeAddress(rowIndex, rowIndex, cellIndex, cellIndex + 3);
        sheet.addMergedRegion(craMatchTeam);
        XSSFCell cellMatchTeam = matchTeam.createCell(cellIndex);
        cellMatchTeam.setCellValue(match.getMatchTeam());
        cellMatchTeam.setCellStyle(styleMap.get("name"));

        double[] minKaiLi8888 = match.getMinKaiLi8888();
        double[] minKaiLiAll = match.getMinKaiLiAll();
        initMinKaiLi(sheet, cellIndex, rowIndex, minKaiLi8888);
        initMinKaiLi(sheet, cellIndex, rowIndex + 2, minKaiLiAll);

        // Write 凯利指数 差值
        XSSFRow kaiLiSubtra = sheet.getRow(rowIndex + 5);
        for (int i = 0; i < minKaiLi8888.length && i < minKaiLiAll.length && i < 3; i++) {
            XSSFCell subtra = kaiLiSubtra.createCell(cellIndex + i);
            subtra.setCellValue((minKaiLi8888[i] - minKaiLiAll[i]) * 100);
            subtra.setCellStyle(styleMap.get("kaiLi"));
        }

        initPeiLv(sheet, rowIndex, cellIndex, match);
    }

    private void initPeiLv(XSSFSheet sheet, int rowIndex, int cellIndex, Match match) {
        double[][] peiLv8888 = match.getPeiLv8888();
        double[][] peiLvAll = match.getPeiLvAll();
        initPeiLv(sheet, rowIndex, cellIndex + 4, peiLv8888);
        initPeiLv(sheet, rowIndex, cellIndex + 7, peiLvAll);
        int peiLvSubtraCellIndex = cellIndex + 10;

        double[][] peiLvSubtra = new double[6][3];
        for (int i = 0; i < 6; i++) {
            for (int j = 0; j < 3; j++) {
                peiLvSubtra[i][j] = peiLv8888[i][j] - peiLvAll[i][j];
            }
        }

        double[] minPeiLvSubtra = { Double.MAX_VALUE, Double.MAX_VALUE, Double.MAX_VALUE };
        for (int i = 0; i < 3; i++) {
            for (int j = 0; j < 6; j++) {
                if (Math.abs(minPeiLvSubtra[i]) > Math.abs(peiLvSubtra[j][i])) {
                    minPeiLvSubtra[i] = peiLvSubtra[j][i];
                }
            }
        }
        for (int i = 0; i < 6; i++) {
            XSSFRow peiLv = sheet.getRow(rowIndex + i);
            for (int j = 0; j < 3; j++) {
                XSSFCell perPeiLv = peiLv.createCell(peiLvSubtraCellIndex + j);
                perPeiLv.setCellValue(peiLvSubtra[i][j]);
                if (peiLvSubtra[i][j] == minPeiLvSubtra[j]) {
                    perPeiLv.setCellStyle(styleMap.get("peiLvSubtraNearZero"));
                } else {
                    perPeiLv.setCellStyle(styleMap.get("kaiLi"));
                }

            }
        }
    }

    private void initPeiLv(XSSFSheet sheet, int rowIndex, int cellIndex, double[][] peiLvs) {

        for (int i = 0; i < 6 && i < peiLvs.length; i++) {
            XSSFRow peiLv = sheet.getRow(rowIndex + i);
            for (int j = 0; j < 3 && j < peiLvs[i].length; j++) {
                XSSFCell perPeiLv = peiLv.createCell(cellIndex + j);
                // sheet.setColumnWidth(cellIndex + j, 1400);
                perPeiLv.setCellValue(peiLvs[i][j]);
                if (i % 2 > 0) {
                    if (peiLvs[i][j] > peiLvs[i - 1][j]) {
                        perPeiLv.setCellStyle(styleMap.get("peiLvMoreThanChuShi"));
                    } else if (peiLvs[i][j] < peiLvs[i - 1][j]) {
                        perPeiLv.setCellStyle(styleMap.get("peiLvLessThanChuShi"));
                    } else {
                        perPeiLv.setCellStyle(styleMap.get("kaiLi"));
                    }
                } else {
                    perPeiLv.setCellStyle(styleMap.get("kaiLi"));
                }
            }
        }
    }

    private void initMinKaiLi(XSSFSheet sheet, int cellIndex, int rowIndex, double[] minKaiLis) {
        // Write 8888 凯利指数
        XSSFRow kaiLi = sheet.getRow(rowIndex + 1);
        for (int i = 0; i < minKaiLis.length && i < 3; i++) {
            CellRangeAddress cra = new CellRangeAddress(rowIndex + 1, rowIndex + 2, cellIndex + i, cellIndex + i);
            sheet.addMergedRegion(cra);
            XSSFCell minKaiLi = kaiLi.createCell(cellIndex + i);
            minKaiLi.setCellValue(minKaiLis[i]);
            minKaiLi.setCellStyle(styleMap.get("kaiLi"));
        }

        CellRangeAddress craCal = new CellRangeAddress(rowIndex + 1, rowIndex + 2, cellIndex + 3, cellIndex + 3);
        sheet.addMergedRegion(craCal);
        XSSFCell cal = kaiLi.createCell(cellIndex + 3, Cell.CELL_TYPE_STRING);
        cal.setCellValue(getCal(minKaiLis));
        cal.setCellStyle(styleMap.get("kaiLi"));

    }

    private String getCal(double[] minKaiLis) {
        String result = "";
        if (minKaiLis[0] < minKaiLis[1] && minKaiLis[1] < minKaiLis[2]) {
            result = "31";
        } else if (minKaiLis[0] <= minKaiLis[2] && minKaiLis[2] <= minKaiLis[1]) {
            result = "30";
        } else if (minKaiLis[1] <= minKaiLis[0] && minKaiLis[0] <= minKaiLis[2]) {
            result = "13";
        } else if (minKaiLis[1] <= minKaiLis[2] && minKaiLis[2] <= minKaiLis[0]) {
            result = "10";
        } else if (minKaiLis[2] <= minKaiLis[0] && minKaiLis[0] <= minKaiLis[1]) {
            result = "03";
        } else if (minKaiLis[2] <= minKaiLis[1] && minKaiLis[1] <= minKaiLis[0]) {
            result = "01";
        }
        return result;
    }

    private static final Set<String> Company_8888 = new HashSet<String>();
    private static final Pattern GAME_PATTERN = Pattern.compile("\"([^\"]+)\"");
    static {
        /**
         * [16, '10BET', 1, 0]<br>
         * [18, '12BET', 0, 0]<br>
         * [281, 'bet 365', 1, 0]<br>
         * [2, 'Betfair', 0, 1]<br>
         * [545, 'SB', 1, 0]<br>
         */
        Company_8888.add("16");
        Company_8888.add("18");
        Company_8888.add("281");
        Company_8888.add("2");
        Company_8888.add("545");
    }

    static class Match {
        private Long id;
        private Map<String, String> params;
        /**
         * "16|60347729|10BET|2.55|3.1|2.55|35.43|29.14|35.43|90.34|2.55|2.95|2.6|35.15|30.38|34.47|89.63|0.90|0.89|0.90|2016,09-1,20,22,53,00|10BET(英国)|1|0"
         * 
         * @1_3 id、、代码<br>
         * @4_6 初盘主和客赔率<br>
         * @7-9 初盘主和客胜率<br>
         * @10 初盘返还率<br>
         * @11_13 即时主和客赔率<br>
         * @14_16 即时主和客胜率<br>
         * @17 即时返还率<br>
         * @18_20 凯利指数
         */
        private Map<String, String[]> gameMap;

        public Match(Long id, Map<String, String> params) {
            this.id = id;
            this.params = params;
            initGameMap();
        }

        public String getMatchTeam() {
            if (params == null) {
                return "";
            }
            String homeTeam = params.get("hometeam_cn") == null ? "" : params.get("hometeam_cn");
            String guestTeam = params.get("guestteam_cn") == null ? "" : params.get("guestteam_cn");
            StringBuilder sb = new StringBuilder();
            sb.append(homeTeam);
            sb.append("(主) VS ");
            sb.append(guestTeam);
            return sb.toString();
        }

        public double[] getMinKaiLi8888() {
            if (gameMap == null) {
                return null;
            }
            double[] minKaiLi8888 = { Double.MAX_VALUE, Double.MAX_VALUE, Double.MAX_VALUE };

            int start = 17;
            for (String companyId : Company_8888) {
                String[] params = gameMap.get(companyId);
                if (params == null) {
                    continue;
                }
                for (int i = 0; i < minKaiLi8888.length; i++) {
                    double kaili = Double.valueOf(params[start + i]);
                    if (kaili < minKaiLi8888[i]) {
                        minKaiLi8888[i] = kaili;
                    }
                }
            }
            return minKaiLi8888;
        }

        public double[] getMinKaiLiAll() {
            if (gameMap == null) {
                return null;
            }
            double[] minKaiLiAll = { Double.MAX_VALUE, Double.MAX_VALUE, Double.MAX_VALUE };
            int start = 17;
            for (String[] params : gameMap.values()) {

                for (int i = 0; i < minKaiLiAll.length; i++) {
                    double kaili = Double.valueOf(params[start + i]);
                    if (kaili < minKaiLiAll[i]) {
                        minKaiLiAll[i] = kaili;
                    }
                }
            }
            return minKaiLiAll;
        }

        public double[][] getPeiLv8888() {
            if (gameMap == null) {
                return null;
            }
            List<String[]> params8888 = new ArrayList<String[]>();
            for (String companyId : Company_8888) {
                String[] params = gameMap.get(companyId);
                if (params == null) {
                    continue;
                }
                params8888.add(params);
            }
            return getPeiLv(params8888);
        }

        public double[][] getPeiLv(Collection<String[]> paramsList) {
            if (params == null) {
                return null;
            }
            double[][] peiLv =
                    {
                            { Double.MIN_VALUE, Double.MIN_VALUE, Double.MIN_VALUE },
                            { Double.MIN_VALUE, Double.MIN_VALUE, Double.MIN_VALUE },
                            { Double.MAX_VALUE, Double.MAX_VALUE, Double.MAX_VALUE },
                            { Double.MAX_VALUE, Double.MAX_VALUE, Double.MAX_VALUE },
                            { 0, 0, 0 },
                            { 0, 0, 0 } };
            int chuPanStart = 3;
            int jiShiStart = 10;
            for (String[] params : paramsList) {
                if (params == null) {
                    continue;
                }
                for (int i = 0; i < peiLv[0].length; i++) {

                    double chuPanPeiLv = NumberUtils.toDouble(params[chuPanStart + i], -1);
                    double jiShiPeiLv = NumberUtils.toDouble(params[jiShiStart + i], chuPanPeiLv);
                    if (chuPanPeiLv > peiLv[0][i]) {
                        peiLv[0][i] = chuPanPeiLv;
                    }
                    if (chuPanPeiLv < peiLv[2][i]) {
                        peiLv[2][i] = chuPanPeiLv;
                    }

                    if (jiShiPeiLv > peiLv[1][i]) {
                        peiLv[1][i] = jiShiPeiLv;
                    }
                    if (jiShiPeiLv < peiLv[3][i]) {
                        peiLv[3][i] = jiShiPeiLv;
                    }
                }
            }
            BigDecimal two = new BigDecimal(2.00);
            for (int i = 0; i < peiLv[0].length; i++) {
                BigDecimal maxCPPLSum = new BigDecimal(peiLv[0][i] + peiLv[2][i]);
                peiLv[4][i] = maxCPPLSum.divide(two, 2, BigDecimal.ROUND_HALF_UP).doubleValue();
                BigDecimal minCPPLSum = new BigDecimal(peiLv[1][i] + peiLv[3][i]);
                peiLv[5][i] = minCPPLSum.divide(two, 2, BigDecimal.ROUND_HALF_UP).doubleValue();
            }
            return peiLv;
        }

        public double[][] getPeiLvAll() {
            if (gameMap == null) {
                return null;
            }
            return getPeiLv(gameMap.values());
        }

        private void initGameMap() {
            if (gameMap != null || params == null || params.get("game") == null) {
                return;
            }
            gameMap = new HashMap<String, String[]>();
            List<String> perCompanyGames = new ArrayList<String>();
            Matcher gameMatcher = GAME_PATTERN.matcher(params.get("game"));
            while (gameMatcher.find()) {
                perCompanyGames.add(gameMatcher.group(1));
            }
            for (String perGame : perCompanyGames) {
                if (perGame != null && !"".equals(perGame.trim())) {
                    String[] perGameParams = perGame.split("\\|");
                    gameMap.put(perGameParams[0], perGameParams);
                }
            }
        }

        public Long getId() {
            return id;
        }

        public Map<String, String> getParams() {
            return params;
        }

    }
}
