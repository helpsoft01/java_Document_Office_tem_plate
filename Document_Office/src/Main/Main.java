package Main;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/*
 * Note
 * ============================
 * Dùng: org.docx4j 	: 2.8.0
 * 		file template	: dotx
 * 		file out 		: doc
 */
public class Main {

	public static void main(String[] args) throws Exception {

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
				.load(new java.io.File(System.getProperty("user.dir") + "/doc3.dotx"));

		List<Map<DataFieldName, String>> data = new ArrayList<Map<DataFieldName, String>>();

		Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
		map.put(new DataFieldName("name"), "Sonvn VŨ Hải Sơn");
		map.put(new DataFieldName("address"), "Kiến An Hải PHòng");
		// map.put(new DataFieldName("Kundenstrasse"), "Bourke Street");

		data.add(map);

		// map = new HashMap<DataFieldName, String>();
		// map.put( new DataFieldName("Kundenname"), "Jason");
		// map.put(new DataFieldName("Kundenstrasse"), "Collins Street");

		// data.add(map);

		// System.out.println(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(),
		// true, true));

		WordprocessingMLPackage output = org.docx4j.model.fields.merge.MailMerger
				.getConsolidatedResultCrude(wordMLPackage, data);

		// System.out.println(XmlUtils.marshaltoString(output.getMainDocumentPart().getJaxbElement(),
		// true, true));

		output.save(new java.io.File(System.getProperty("user.dir") + "/doc1-OUT.doc"));

	};
}