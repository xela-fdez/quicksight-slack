package com.xela.quicksightslack;

import com.slack.api.Slack;
import com.slack.api.methods.request.chat.ChatPostMessageRequest;
import com.slack.api.methods.response.chat.ChatPostMessageResponse;
import java.io.FileInputStream;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.json.JSONArray;

public class SlackUpload {
	private String apiToken;
	private String channel;
	private Slack slack;
	
	public SlackUpload(String apiToken, String channel) {
		this.apiToken = apiToken;
		this.channel = channel;
		slack = Slack.getInstance();
	}
	
	public void upload(String fileName, int sheetNumber) throws Exception {

		FileInputStream inputStream = new FileInputStream(fileName);
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		
		System.out.println(new ExcelToJSONArray(workbook, sheet).getJSONArray());
		JSONArray table = (new ExcelToJSONArray(workbook, sheet)).getJSONArray();

		StringBuilder sbTable = JSONArrayToStringBuilder(table);

		ChatPostMessageRequest request = ChatPostMessageRequest 													
				.builder().channel(channel).text(codeBlock(sbTable.toString())).build();

		ChatPostMessageResponse response = slack.methods(apiToken).chatPostMessage(request);
		if (response.isOk()) {
			System.out.println("Message sent successfully");
		} else {
			System.out.println("Message send failed: " + response.getError());
		}
	}
	
	public void uploadBacklogs(String fileName, int sheetNumber) throws Exception {
		
		
		DecimalFormat df = new DecimalFormat("0.00");

		FileInputStream inputStream = new FileInputStream(fileName);
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(sheetNumber);

		System.out.println(new ExcelToJSONArray(workbook, sheet).getJSONArray());
		JSONArray table = (new ExcelToJSONArray(workbook, sheet)).getJSONArray();
		JSONArray breach = (new JSONArray()).put(table.getJSONArray(0));
		JSONArray highAgeing = (new JSONArray()).put(table.getJSONArray(0));
		String mp = null;
		int i;
		for (i = 0; i < table.length(); i++) {
			if (!table.getJSONArray(i).get(0).equals(""))
				mp = table.getJSONArray(i).get(0).toString();
			try {
				double ageing = Double.parseDouble(table.getJSONArray(i).get(5).toString());
				if (table.getJSONArray(i).getString(2).equals("SPONSORED_PRODUCTS")) {
					if (ageing > 24.0D) {
						breach.put(table.getJSONArray(i));
						breach.getJSONArray(breach.length() - 1).put(0, mp);
					} else if (ageing > 16.0D) {
						highAgeing.put(table.getJSONArray(i));
						highAgeing.getJSONArray(highAgeing.length() - 1).put(0, mp);
					}
				} else if (ageing > 12.0D) {
					breach.put(table.getJSONArray(i));
					breach.getJSONArray(breach.length() - 1).put(0, mp);
				} else if (ageing > 8.0D) {
					highAgeing.put(table.getJSONArray(i));
					highAgeing.getJSONArray(highAgeing.length() - 1).put(0, mp);
				}
			} catch (NumberFormatException numberFormatException) {
			}
		}
		for (i = 0; i < table.length(); i++) {
			try {
				table.getJSONArray(i).put(5, df.format(Double.parseDouble(table.getJSONArray(i).get(5).toString())));
			} catch (NumberFormatException numberFormatException) {
			}
		}
		for (i = 0; i < breach.length(); i++) {
			try {
				breach.getJSONArray(i).put(5, df.format(Double.parseDouble(breach.getJSONArray(i).get(5).toString())));
			} catch (NumberFormatException numberFormatException) {
			}
		}
		for (i = 0; i < highAgeing.length(); i++) {
			try {
				highAgeing.getJSONArray(i).put(5, df.format(Double.parseDouble(highAgeing.getJSONArray(i).get(5).toString())));
			} catch (NumberFormatException numberFormatException) {
			}
		}
		StringBuilder sbTable = JSONArrayToStringBuilder(table);
		StringBuilder sbBreach = JSONArrayToStringBuilder(breach);
		StringBuilder sbHighAgeing = JSONArrayToStringBuilder(highAgeing);
		
		ChatPostMessageRequest request = ChatPostMessageRequest														
				.builder().channel(channel).text(codeBlock(sbTable.toString()) + "\nBreach:\n" + 
						codeBlock(sbBreach.toString())+ "\nHigh Ageing:\n" + codeBlock(sbHighAgeing.toString()))
				.build();

		ChatPostMessageResponse response = slack.methods(apiToken).chatPostMessage(request);
		if (response.isOk()) {
			System.out.println("Message sent successfully");
		} else {
			System.out.println("Message send failed: " + response.getError());
		}
	}
	
	

	private static StringBuilder JSONArrayToStringBuilder(JSONArray table) {
		int maxLength = 0;
		for (int i = 0; i < table.length(); i++) {
			if (table.getJSONArray(i).length() > maxLength)
				maxLength = table.getJSONArray(i).length();
		}

		int[] lengths = new int[maxLength];
		int i;
		for (i = 0; i < lengths.length; i++)
			lengths[i] = 0;
		for (i = 0; i < table.length(); i++) {
			for (int j = 0; j < table.getJSONArray(i).length(); j++) {
				if (table.getJSONArray(i).get(j).toString().length() > lengths[j])
					lengths[j] = table.getJSONArray(i).get(j).toString().length();
			}
		}
		StringBuilder sb = new StringBuilder();
		for (int j = 0; j < table.length(); j++) {
			sb.append("|");
			for (int k = 0; k < table.getJSONArray(j).length(); k++) {
				sb.append(table.getJSONArray(j).get(k).toString());
				for (int m = table.getJSONArray(j).get(k).toString().length(); m < lengths[k]; m++)
					sb.append(" ");
				sb.append("|");
			}
			sb.append("\n");
		}
		return sb;
	}

	private static String codeBlock(String text) {
		StringBuilder sb = new StringBuilder();
		StringBuilder aux = new StringBuilder(text);
		int previousInserted = 0;

//		"```" + sbTable.toString() + "```";
		
		while(aux.length()>0) {
			sb.append("```");
			while (aux.length()>0 && (sb.length() + aux.substring(0, aux.indexOf("\n")).length()) < 3996+previousInserted) {
				sb.append(aux.substring(0, aux.indexOf("\n")+1));
				aux.delete(0, aux.indexOf("\n")+1);
			}
			sb.append("```\n");
			previousInserted=sb.length();
		}

		return sb.toString();
	}
}
