# POI-word-
POI操作word文档

## springmvc返回代码创建的简单word
```java
@RequestMapping("/export1/")
//	@RequiresPermissions(value={"schedule:export"})
	public ResponseEntity<byte[]> export1() throws Exception {
		XWPFDocument doc = new XWPFDocument();// 创建Word文件
		XWPFParagraph p = doc.createParagraph();// 新建一个段落
		p.setAlignment(ParagraphAlignment.CENTER);// 设置段落的对齐方式
		XWPFRun r = p.createRun();//创建段落文本
        
    r.setText("（2018）第四次局长办公室会议通知");
		r.setBold(true);//设置为粗体
		r.setFontSize(20);//设置字号
		
		//基本信息表格
		XWPFTable infoTable = doc.createTable();
//        infoTable.getCTTbl().getTblPr().unsetTblBorders();//去表格边框
    //表格第1行（时间）
    XWPFTableRow infoTableRow1 = infoTable.getRow(0);
    infoTableRow1.setHeight(300);
    //设置单元格宽度
		XWPFParagraph p1 = infoTableRow1.getCell(0).addParagraph();
		CTTcPr tcpr = infoTableRow1.getCell(0).getCTTc().addNewTcPr();
		CTTblWidth cellWidth = tcpr.addNewTcW();
		cellWidth.setType(STTblWidth.DXA);
		cellWidth.setW(BigInteger.valueOf(1200));
		
    p1.setAlignment(ParagraphAlignment.DISTRIBUTE);// 设置段落两边散开
    p1.setIndentationRight(100);//设置右侧缩进
    p1.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r1 = p1.createRun();
		r1.setText("时间");
		r1.setBold(true);//设置为粗体
		r1.setFontSize(16);//设置字号
		infoTableRow1.getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell中文字对齐方式
		
		CTTblWidth infoTableWidth1 = infoTable.getCTTbl().addNewTblPr().addNewTblW();
    infoTableWidth1.setType(STTblWidth.DXA);
    infoTableWidth1.setW(BigInteger.valueOf(8500));
    XWPFTableCell cell1 = infoTableRow1.addNewTableCell();
    XWPFParagraph p1_1 = cell1.addParagraph();// 新建一个段落
    p1_1.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p1_1.setFontAlignment(2);
    p1_1.setWordWrapped(true);
    XWPFRun r1_1 = p1_1.createRun();
        
		r1_1.setText("设置段落居中");
		r1_1.setFontSize(14);//设置字号
		cell1.setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell文字对齐方式
       
		//表格第2行（地点）
    XWPFTableRow infoTableRow2 = infoTable.createRow();
    infoTableRow2.setHeight(300);
    //设置单元格宽度
		XWPFParagraph p2 = infoTableRow2.getCell(0).addParagraph();
		CTTcPr tcpr2 = infoTableRow2.getCell(0).getCTTc().addNewTcPr();
		CTTblWidth cellWidth2 = tcpr2.addNewTcW();
		cellWidth2.setType(STTblWidth.DXA);
		cellWidth2.setW(BigInteger.valueOf(1200));
		
		
    p2.setAlignment(ParagraphAlignment.DISTRIBUTE);// 设置段落两边散开
    p2.setIndentationRight(100);//设置右侧缩进
    p2.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r2 = p2.createRun();
		r2.setText("地点");
		r2.setBold(true);//设置为粗体
		r2.setFontSize(16);//设置字号
		infoTableRow2.getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell中文字对齐方式
		
		CTTblWidth infoTableWidth2 = infoTable.getCTTbl().addNewTblPr().addNewTblW();
    infoTableWidth2.setType(STTblWidth.DXA);
    infoTableWidth2.setW(BigInteger.valueOf(8500));
    XWPFTableCell cell2 = infoTableRow2.getCell(1);//addNewTableCell();
    XWPFParagraph p2_2 = cell2.addParagraph();// 新建一个段落
    p2_2.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p2_2.setFontAlignment(2);
    p2_2.setWordWrapped(true);
    XWPFRun r2_2 = p2_2.createRun();

    r2_2.setText("设置段落居中");
    r2_2.setFontSize(14);//设置字号
		cell2.setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell文字对齐方式
		
		//表格第3行（主持人）
    XWPFTableRow infoTableRow3 = infoTable.createRow();
    infoTableRow3.setHeight(300);
    //设置单元格宽度
		XWPFParagraph p3 = infoTableRow3.getCell(0).addParagraph();
		CTTcPr tcpr3 = infoTableRow3.getCell(0).getCTTc().addNewTcPr();
		CTTblWidth cellWidth3 = tcpr3.addNewTcW();
		cellWidth3.setType(STTblWidth.DXA);
		cellWidth3.setW(BigInteger.valueOf(1200));

    p3.setAlignment(ParagraphAlignment.DISTRIBUTE);// 设置段落两边散开
    p3.setIndentationRight(100);//设置右侧缩进
    p3.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r3 = p3.createRun();
		r3.setText("主持人");
		r3.setBold(true);//设置为粗体
		r3.setFontSize(16);//设置字号
		infoTableRow3.getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell中文字对齐方式
		
		CTTblWidth infoTableWidth3 = infoTable.getCTTbl().addNewTblPr().addNewTblW();
    infoTableWidth3.setType(STTblWidth.DXA);
    infoTableWidth3.setW(BigInteger.valueOf(8500));
    XWPFTableCell cell3 = infoTableRow3.getCell(1);//addNewTableCell();
    XWPFParagraph p3_3 = cell3.addParagraph();// 新建一个段落
    p3_3.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p3_3.setFontAlignment(2);
    p3_3.setWordWrapped(true);
    XWPFRun r3_3 = p3_3.createRun();

    r3_3.setText("设置段");
    r3_3.setFontSize(14);//设置字号
		cell3.setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell文字对齐方式
		
		
		//表格第4行（会议议题）
    XWPFTableRow infoTableRow4 = infoTable.createRow();
    infoTableRow4.setHeight(4500);
    //设置单元格宽度
		CTTcPr tcpr4 = infoTableRow4.getCell(0).getCTTc().addNewTcPr();
		CTTblWidth cellWidth4 = tcpr4.addNewTcW();
		cellWidth4.setType(STTblWidth.DXA);
		cellWidth4.setW(BigInteger.valueOf(1200));
		
		infoTableRow4.getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell中文字对齐方式
		
		XWPFParagraph p4 = infoTableRow4.getCell(0).addParagraph();
    p4.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p4.setIndentationRight(100);//设置右侧缩进
    p4.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r4 = p4.createRun();
    r4.setText("会");
    r4.setBold(true);//设置为粗体
    r4.setFontSize(16);//设置字号

    XWPFParagraph p4_1 = infoTableRow4.getCell(0).addParagraph();
    p4_1.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p4_1.setIndentationRight(100);//设置右侧缩进
    p4_1.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r4_1 = p4_1.createRun();
    r4_1.setText("议");
    r4_1.setBold(true);//设置为粗体
    r4_1.setFontSize(16);//设置字号

    XWPFParagraph p4_2 = infoTableRow4.getCell(0).addParagraph();
    p4_2.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p4_2.setIndentationRight(100);//设置右侧缩进
    p4_2.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r4_2 = p4_2.createRun();
    r4_2.setText("议");
    r4_2.setBold(true);//设置为粗体
    r4_2.setFontSize(16);//设置字号

    XWPFParagraph p4_3 = infoTableRow4.getCell(0).addParagraph();
    p4_3.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p4_3.setIndentationRight(100);//设置右侧缩进
    p4_3.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r4_3 = p4_3.createRun();
    r4_3.setText("题");
    r4_3.setBold(true);//设置为粗体
    r4_3.setFontSize(16);//设置字号

		
		CTTblWidth infoTableWidth4 = infoTable.getCTTbl().addNewTblPr().addNewTblW();
    infoTableWidth4.setType(STTblWidth.DXA);
    infoTableWidth4.setW(BigInteger.valueOf(8500));
    XWPFTableCell cell4 = infoTableRow4.getCell(1);//addNewTableCell();
    XWPFParagraph p4_4 = cell4.addParagraph();// 新建一个段落
    p4_4.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p4_4.setFontAlignment(2);
    p4_4.setWordWrapped(true);
    XWPFRun r4_4 = p4_4.createRun();

    r4_4.setText("设置段");
    r4_4.setFontSize(14);//设置字号
		cell4.setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell文字对齐方式
		
		
		//表格第5行（参见人员）
    XWPFTableRow infoTableRow5 = infoTable.createRow();
    infoTableRow5.setHeight(4500);
    //设置单元格宽度
		CTTcPr tcpr5 = infoTableRow5.getCell(0).getCTTc().addNewTcPr();
		CTTblWidth cellWidth5 = tcpr5.addNewTcW();
		cellWidth5.setType(STTblWidth.DXA);
		cellWidth5.setW(BigInteger.valueOf(1200));
		
		infoTableRow5.getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell中文字对齐方式
		
		XWPFParagraph p5 = infoTableRow5.getCell(0).addParagraph();
		p5.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
		p5.setIndentationRight(100);//设置右侧缩进
		p5.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r5 = p5.createRun();
    r5.setText("参");
    r5.setBold(true);//设置为粗体
    r5.setFontSize(16);//设置字号

    XWPFParagraph p5_1 = infoTableRow5.getCell(0).addParagraph();
    p5_1.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p5_1.setIndentationRight(100);//设置右侧缩进
    p5_1.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r5_1 = p5_1.createRun();
    r5_1.setText("加");
    r5_1.setBold(true);//设置为粗体
    r5_1.setFontSize(16);//设置字号

    XWPFParagraph p5_2 = infoTableRow5.getCell(0).addParagraph();
    p5_2.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p5_2.setIndentationRight(100);//设置右侧缩进
    p5_2.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r5_2 = p5_2.createRun();
    r5_2.setText("人");
    r5_2.setBold(true);//设置为粗体
    r5_2.setFontSize(16);//设置字号

    XWPFParagraph p5_3 = infoTableRow5.getCell(0).addParagraph();
    p5_3.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p5_3.setIndentationRight(100);//设置右侧缩进
    p5_3.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r5_3 = p5_3.createRun();
    r5_3.setText("员");
    r5_3.setBold(true);//设置为粗体
    r5_3.setFontSize(16);//设置字号


    CTTblWidth infoTableWidth5 = infoTable.getCTTbl().addNewTblPr().addNewTblW();
    infoTableWidth5.setType(STTblWidth.DXA);
    infoTableWidth5.setW(BigInteger.valueOf(8500));
    XWPFTableCell cell5 = infoTableRow5.getCell(1);//addNewTableCell();
    XWPFParagraph p5_4 = cell5.addParagraph();// 新建一个段落
    p5_4.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p5_4.setFontAlignment(2);
    p5_4.setWordWrapped(true);
    XWPFRun r5_4 = p5_4.createRun();

    r5_4.setText("设置段");
    r5_4.setFontSize(14);//设置字号
		cell5.setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell文字对齐方式
		
		
		//表格第6行（签发）
    XWPFTableRow infoTableRow6 = infoTable.createRow();
    infoTableRow6.setHeight(1500);
    //设置单元格宽度
		CTTcPr tcpr6 = infoTableRow6.getCell(0).getCTTc().addNewTcPr();
		CTTblWidth cellWidth6 = tcpr6.addNewTcW();
		cellWidth6.setType(STTblWidth.DXA);
		cellWidth6.setW(BigInteger.valueOf(600));
		
		infoTableRow6.getCell(0).setVerticalAlignment(XWPFVertAlign.CENTER);// 设置cell中文字对齐方式
		
		XWPFParagraph p6 = infoTableRow6.getCell(0).addParagraph();
		p6.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
		p6.setIndentationRight(100);//设置右侧缩进
		p6.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r6 = p6.createRun();
    r6.setText("签");
    r6.setBold(true);//设置为粗体
    r6.setFontSize(16);//设置字号

    XWPFParagraph p6_1 = infoTableRow6.getCell(0).addParagraph();
    p6_1.setAlignment(ParagraphAlignment.CENTER);// 设置段落居中
    p6_1.setIndentationRight(100);//设置右侧缩进
    p6_1.setIndentationLeft(100);//设置左侧缩进
    XWPFRun r6_1 = p6_1.createRun();
    r6_1.setText("发");
    r6_1.setBold(true);//设置为粗体
    r6_1.setFontSize(16);//设置字号

    XWPFParagraph p7 = doc.createParagraph();// 新建一个段落
    p7.setAlignment(ParagraphAlignment.CENTER);// 设置段落的对齐方式
		XWPFRun r7 = p7.createRun();//创建段落文本
        
    r7.setText("注");
		r7.setBold(true);//设置为粗体
		r7.setFontSize(10);//设置字号
		
		XWPFRun r7_1 = p7.createRun();//创建段落文本
		r7_1.setText("：请参会人员严格遵守保密协议，不要携带手机和装有无线发射功能的设备及非保密笔记本电脑带入会场");
//		r7_1.setBold(true);//设置为粗体
		r7_1.setFontSize(10);//设置字号
		
		FileOutputStream out = new FileOutputStream("F:\\POI\\sample.docx");
		System.out.println("66666666666");
		
		
		ByteArrayOutputStream byteArrayOutputStream = null;
		ResponseEntity<byte[]> responseEntity = null;
		// 将文件写进字节流中
		byteArrayOutputStream = new ByteArrayOutputStream();
		doc.write(byteArrayOutputStream);
		byte[] byteArray = byteArrayOutputStream.toByteArray();
		
		// 返回设置
		HttpHeaders headers = new HttpHeaders();
		headers = new HttpHeaders();
		MediaType mediaType = new MediaType(MediaType.APPLICATION_OCTET_STREAM, Charset.forName("utf-8"));
		headers.setContentType(mediaType);
		String fileName = URLEncoder.encode("会议通知.docx", "utf-8");
		headers.setContentDispositionFormData("attachment", fileName);

		// 返回文件
		responseEntity = new ResponseEntity<byte[]>(byteArray, headers,
				HttpStatus.CREATED);
		
//		doc.write(out);
		out.close();
		doc.close();
		return responseEntity;
		
	}
  ```
  ## springmvc返回模板word
  ```java
  @RequestMapping("/export/{current}")
//	@RequiresPermissions(value={"schedule:export"})
	public ResponseEntity<byte[]> export(@PathVariable String current) {
		URL url = this.getClass().getClassLoader().getResource("scheduleExportTemplate.doc");
		String tmpFile =  url.getPath();
		StringBuffer currentSub = new StringBuffer().append(current.substring(0, 7));
		List<Schedule> list = new ArrayList<>();
		// 当月第一天
		Date fristDay = DateUtil.parseStrToDate(current, "yyyy-MM-dd");
		Instant instant = fristDay.toInstant();
		ZoneId zoneId = ZoneId.systemDefault();
		LocalDate localDate = instant.atZone(zoneId).toLocalDate();
		LocalDate localDateLastDay = localDate.with(TemporalAdjusters.lastDayOfMonth());
		// 当月最后一天
		ZonedDateTime zdt = localDateLastDay.atStartOfDay(zoneId);
		Date lastDay = Date.from(zdt.toInstant());
		Schedule schedule = new Schedule();
		schedule.getMap().put("fristDay", fristDay);
		schedule.getMap().put("lastDay", lastDay);
		schedule.getMap().put("orderBy", "scheduled_time");
		schedule.getMap().put("sort", "asc");
		list = this.scheduleService.selectEntityList(schedule);

		Map<String, String> datas = new HashMap<String, String>();
		datas.put("current", current);
		for (int i = 0; i < list.size(); i++) {
			int j = i + 1;
			Schedule schedule1 = list.get(i);
			String start = DateUtil.parseDateToStr(schedule1.getScheduledTime(), "yyyy-MM-dd");
			String title = schedule1.getMap().get("person_name").toString();
			datas.put("time" + j, start);
			datas.put("name" + j, title);
		}
		FileInputStream tempFileInputStream = null;
		HWPFDocument document = null;
		ByteArrayOutputStream byteArrayOutputStream = null;
		ResponseEntity<byte[]> responseEntity = null;
		try {
			tempFileInputStream = new FileInputStream(tmpFile);
			document = new HWPFDocument(tempFileInputStream);
			// 读取文本内容
			Range bodyRange = document.getRange();
			// 替换内容
			for (Map.Entry<String, String> entry : datas.entrySet()) {
				bodyRange.replaceText("{{" + entry.getKey() + "}}", entry.getValue());
			}
			// 将文件写进字节流中
			byteArrayOutputStream = new ByteArrayOutputStream();
			document.write(byteArrayOutputStream);
			byte[] byteArray = byteArrayOutputStream.toByteArray();
			
			// 返回设置
			HttpHeaders headers = new HttpHeaders();
			headers = new HttpHeaders();
			MediaType mediaType = new MediaType(MediaType.APPLICATION_OCTET_STREAM, Charset.forName("utf-8"));
			headers.setContentType(mediaType);
			String fileName = URLEncoder.encode(currentSub.append("值班信息.doc").toString(), "utf-8");
			headers.setContentDispositionFormData("attachment", fileName);

			// 返回文件
			responseEntity = new ResponseEntity<byte[]>(byteArray, headers,
					HttpStatus.CREATED);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				tempFileInputStream.close();
				document.close();
				byteArrayOutputStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return responseEntity;
	}
  ```
