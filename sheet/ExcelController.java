
public class EntPortraitExcelController {


    //导入表格
    //获取表头
    //校验表头
    //获取数据
    //校验数据
    //保存数据

    public ResultModel parseDataExcel(String importId, String year, String dataName) {
        sysMongoService.saveOpLog(SysConstants.LEVEL_HIGH, "导入企业画像", year + "年" + dataName + "数据");
        RestResultBuilder builder = RestResultBuilder.builder();
        try {
            List<Sheet> sheets = getSheets(importId);
            List<String> logList = new ArrayList<>();
            String msg = null;
            if (!sheets.isEmpty()) {
                Sheet sheet = sheets.get(0);
                // 校验sheet是否合法
                if (sheet == null) {
                    return null;
                }
                checkRepeatData(sheet, sheet.getFirstRowNum(), sheet.getFirstRowNum() + 2);
                String headerCheckName = null;
                checkExcelHead(headerCheckName, sheet.getRow(sheet.getFirstRowNum()));
                List<EntPortraitPo> list = parseExcel(sheet);
            }
            if (!logList.isEmpty()) {
                logger.error("录入失败数据:" + StringUtils.join(logList, ","));
                msg = "录入失败数据:" + StringUtils.join(logList, ",");
            }
            return builder.success(msg).build();
        } catch (PoiException e) {
            String err = e.getErrorMsg();
            return builder.message(err).build();
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
            String err = "解析报错";
            return builder.message(err).build();
        }
    }

    /**
     * 获取sheet
     *
     * @param importId
     * @return
     */
    private List<Sheet> getSheets(String importId) {
        return ExcelUtil.readExcl(path);
    }

    /**
     * 解析excel
     *
     * @param sheet
     * @return
     */
    private ArrayList<EntPortraitPo> parseExcel(Sheet sheet) throws Exception {
        // 校验sheet是否合法
        if (sheet == null) {
            return null;
        }
        // 获取第一行数据
        int firstRowNum = sheet.getFirstRowNum();
        Row firstRow = sheet.getRow(firstRowNum);
        if (null == firstRow) {
            logger.error("解析Excel失败，在第一行没有读取到任何数据！");
        }
        // 解析每一行的数据，构造数据对象
        //第二行开始解析
        int rowStart = firstRowNum + 2;
        int rowEnd = sheet.getPhysicalNumberOfRows();
        ArrayList<List<String>> header = ExcelUtil.makeHeader(sheet.getRow(0), sheet.getRow(1));
        ArrayList<EntPortraitPo> resultDataList = new ArrayList<>();
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (null == row) {
                continue;
            }
            resultDataList.addAll(makeData(row, header));
        }
        return resultDataList;
    }


    /**
     * 组合数据
     *
     * @param row
     * @param header
     * @return
     */
    private static ArrayList<EntPortraitPo> makeData(Row row, ArrayList<List<String>> header) throws Exception {
        HashMap<String, EntPortraitPo> monthDataMap = new HashMap<>();
        ArrayList<EntPortraitPo> dataList = new ArrayList<>();
        String entName = "";
        String uniformCreditCode = "";
        for (int i = 0; i < row.getLastCellNum(); i++) {
            //第一行
            List<String> firstRow = header.get(0);
            //第二行
            List<String> secondRow = header.get(1);
            String thisRowData = row.getCell(i) != null ? row.getCell(i).toString() : "";
            //第二行没有数据的话，表示数据是企业基础数据，有的话代表是录入的关联数据
            if (secondRow.get(i).equals("")) {
                switch (firstRow.get(i)) {
                    case "企业名称":
                        entName = row.getCell(i).toString();
                        break;
                    case "企业统一社会信用代码":
                        uniformCreditCode = row.getCell(i).toString();
                        break;
                }
            } else {
                String[] secondRowList = secondRow.get(i).split("月份");
                String month = secondRowList[0];
                EntPortraitPo monthEntPortraitPo = monthDataMap.get(month);
                if (monthEntPortraitPo != null) {
                    setExcelDataToEntPortrait(firstRow.get(i).split("\n")[0], thisRowData, monthEntPortraitPo);
                } else {
                    EntPortraitPo entPortraitPo = new EntPortraitPo();
                    entPortraitPo.setEntName(entName);
                    entPortraitPo.setUniformCreditCode(uniformCreditCode);
                    entPortraitPo.setMonth(month);
                    monthDataMap.put(month, entPortraitPo);
                    setExcelDataToEntPortrait(firstRow.get(i).split("\n")[0], thisRowData, entPortraitPo);
                }
            }
        }
        for (String key : monthDataMap.keySet()) {
            dataList.add(monthDataMap.get(key));
        }
        return dataList;
    }

    public ResultModel parseTaxRateExcel(String importId, String year, String month) {
        RestResultBuilder builder = RestResultBuilder.builder();
        try {
            List<Sheet> sheets = getSheets(importId);
            List<String> logList = new ArrayList<>();
            if (!sheets.isEmpty()) {
                Sheet sheet = sheets.get(0);
                // 校验sheet是否合法
                if (sheet == null) {
                    return null;
                }
                checkRepeatData(sheet, sheet.getFirstRowNum() + 1, sheet.getFirstRowNum() + 2);
                checkExcelHead("zb", sheet.getRow(sheet.getFirstRowNum() + 1));
                // 获取第一行数据
                int firstRowNum = sheet.getFirstRowNum();
                Row firstRow = sheet.getRow(firstRowNum);
                if (null == firstRow) {
                    logger.error("解析Excel失败，在第一行没有读取到任何数据！");
                }
                // 解析每一行的数据，构造数据对象
                //第二行开始解析
                int rowStart = firstRowNum + 2;
                int rowEnd = sheet.getPhysicalNumberOfRows();
                Row header = sheet.getRow(1);
                for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    if (null == row) {
                        continue;
                    }
                    if (row.getCell(1) == null || row.getCell(1).toString().equals("")) {
                        continue;
                    }
                    EntTaxChangeRateData entTaxChangeRateData = new EntTaxChangeRateData();
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        switch (header.getCell(i).toString()) {
                            case "":
                                entTaxChangeRateData.setEntName(row.getCell(i).toString());
                                break;
                            case "":
                                entTaxChangeRateData.setUniformCreditCode(row.getCell(i).toString());
                                break;
                            case "":
                                entTaxChangeRateData.setFixIndicators(getRate(row.getCell(i).toString()));
                                break;
                        }
                    }
                    resultDataList.add(entTaxChangeRateData);
                }
                logList = entPortraitExcelService.saveTaxChangeRate(resultDataList, month, year);
            }
            return builder.success(StringUtils.join(logList, ",")).build();
        } catch (PoiException e) {
            String err = e.getErrorMsg();
            return builder.message(err).build();
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
            String err = "解析报错";
            return builder.message(err).build();
        }
    }


    /**
     * 检查重复数据
     *
     * @param sheet
     * @param startNum
     * @param headLinkNum
     * @return
     */
    private void checkRepeatData(Sheet sheet, int headLinkNum, int startNum) {
        Row headRow = sheet.getRow(headLinkNum);
        int entCodeNum = 99;
        int entNameNum = 99;
        for (int i = 0; i < headRow.getLastCellNum(); i++) {
            if (headRow.getCell(i) != null && headRow.getCell(i).toString().equals("")) {
                entCodeNum = i;
            }
            if (headRow.getCell(i) != null && headRow.getCell(i).toString().equals("")) {
                entNameNum = i;
            }
        }
        if (entCodeNum == 99) {
            throw new PoiException("检查表头有没有：");
        }
        if (entNameNum == 99) {
            throw new PoiException("检查表头有没有：");
        }
        int rowEnd = sheet.getPhysicalNumberOfRows();
        ArrayList<String> list = new ArrayList<>();
        for (int i = startNum; i < rowEnd; i++) {
            if (sheet.getRow(i) == null) {
                continue;
            }
            Cell cell = sheet.getRow(i).getCell(entCodeNum);
            Cell nameCell = sheet.getRow(i).getCell(entNameNum);
            if (cell == null && nameCell == null) {
                continue;
            }
            if (cell == null || cell.toString().equals("")) {
                String msg = "录入的" + nameCell.toString() + "";
                throw new PoiException(msg);
            }
            list.add(cell.toString());
        }
        HashSet<String> set = new HashSet<>(list);
        if (list.size() != set.size()) {
            String msg = "录入的数据有重复的";
            throw new PoiException(msg);
        }
    }

    /**
     * 检查excel的表头有没有符合录入的内容
     */
    private void checkExcelHead(String name, Row headerRow) {
        List<String> list = headDef.getContextList(name);
        int size = list.size();
        int excelHeadNum = 0;
        short lastCellNum = headerRow.getLastCellNum();
        for (String headName : list) {
            for (int i = 0; i < lastCellNum; i++) {
                if (headerRow.getCell(i) != null && headerRow.getCell(i).toString().split("\n")[0].equals(headName)) {
                    excelHeadNum++;
                }
            }
        }
        if (size != excelHeadNum) {
            throw new PoiException("检查是否含有要录入的表头:" + String.join(",", list));
        }
    }
}
