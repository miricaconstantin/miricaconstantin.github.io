import java.io.*;
import java.math.RoundingMode;
import java.net.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.security.*;
import java.sql.*;
import java.sql.Timestamp;
import java.text.*;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Date;
import java.util.regex.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import java.time.Instant;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.xml.sax.ErrorHandler;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;
import javax.xml.XMLConstants;
import javax.xml.transform.stream.StreamSource;
import javax.xml.validation.Schema;
import javax.xml.validation.SchemaFactory;
import javax.xml.validation.Validator;
import java.math.BigDecimal;


public class XlsToXmlConverter
{
    static String JDBC_DRIVER;
    static String DB_URL;
    static String USER;
    static String PASS;
    static boolean hasitemsA3 ;
    static boolean hasitemsA2 ;
    static boolean hasitemsA1 ;
    static boolean hasitemsA1a ;
    static boolean hasitemsA4 ;
    static String connectionString ="Unknown";
    static String SenderMSGID;
    static boolean headerWritten = false;
    static Set<String> openedCurrencies = new HashSet<>();
    static Set<String> writtenSpecificAppendices = new HashSet<>();

    static Set<String> buyerNos = new HashSet<>();
    static Set<String> buyerUIDs = new HashSet<>();

    static Set<String> form11BuyerUIDs = new HashSet<>();

    static Set<String> form12Unique = new HashSet<>();

    static Set<String> form2BuyerUIDs = new HashSet<>();
    Set<String> allowed = new HashSet<>(Arrays.asList(
            "A","B","C","D","E","F","G","H"
    ));
    static List<String> validationErrors = new ArrayList<>();
    static boolean hasErrors =  false;



    static String path;
    static StringWriter sw = new StringWriter();

    static String sqlQuery;
    static String sqlQuery_seq;
    static ResultSet resultSet;
    static ResultSet SenderMessageId;

    static String xlsFileId;
    static String xlsFileName;
    static String xlsFileIgnoredColors;

    static String reportsGroup;
    static String refDate;
    static String consolidatedFlag;
    static String pathToJar;
    static String reportingEntity;
    static String documentName;
    static String newLine;
    static String validationEnabled;
    static String evaluationFormulasEnabled;
    static String xsdSchema;

    static Clob clob;
    static int refDay;
    static int refMonth;
    static int refYear;
    static FormulaEvaluator evaluator;

    static String refDateSql;

    //adaugat pt SPD270
    static String[][] SPD270_Contacts = new String[10][10];
    static String[][] SPD279_Contacts = new String[10][10];
    static int nrContacte = 0;
    static ResultSet resultSetContacts;
    static PreparedStatement preparedStmt;
    static int exId;

    static Connection connection;
    static Statement statement;
    static Statement statement_seq;

    static Workbook workbook;
    static Sheet sheet;
    static String xmlFileName;
    static String xmlFileNameIMV;
    static String xmlFileNameRM;
    static PrintWriter xmlWriter;

    //Adaugat pentru SBP unde sunt 3 fisiere
    static PrintWriter xmlWriterIMV;
    static PrintWriter xmlWriterRM;

    static String currentSheetName;
    static String reportName;
    static String xmlRowNamePattern;
    static String xmlColumnNamePattern;
    static String xmlPattern;

    static String tableRange;
    static String tableRowCodeColumn;
    static String tableColumnCodeRow;
    static String tableB790Column;
    static String tableGroupfieldColumn;
    static String tableRankColumn;
    static String tableB015Value;
    static String tableI056Cell;
    static String valueOfI056;
    static String tableB272Cell;
    static String valueOfB272;
    static String tableB271Cell;
    static String valueOfB271;
    static String tableC007Cell;
    static String valueOfC007;
    static String tableC200Cell;
    static String valueOfC200;
    static String tableC010Cell;
    static String valueOfC010;
    static String startColumn;
    static String endColumn;
    static Integer startRow;
    static Integer endRow;

    static DataFormatter dataFormatter = new DataFormatter(Locale.US);
    static DecimalFormat decimalFormat = new DecimalFormat("0", DecimalFormatSymbols.getInstance(Locale.ENGLISH));

    static Pattern columnPattern = Pattern.compile("^\\D+");
    static Pattern rowPattern = Pattern.compile("\\d+$");
    public static class VerificariBalanceCode {

        // definim toate listele
        public static final Set<String> coduriIncasariSauPlati = new HashSet<>(Arrays.asList(
                "F50", "H11", "H15", "H20", "H10", "F60", "H00", "D10", "G00", "G05", "G10", "G20",
                "G50", "H33", "E03", "H41", "H42", "H43", "A30", "B21", "B23", "B22", "B01", "B03", "B02",
                "B31", "B33", "B32", "D00", "D20", "B41", "H70", "H80", "H90", "H92", "H60", "K39", "K79",
                "K82", "L17", "L86", "L35", "L18", "L26", "L29", "L34", "M10", "M25", "M50", "P41", "P42",
                "C05", "L12", "L11", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "REC1"
        ));

        public static final Set<String> coduriDoarPlati = new HashSet<>(Arrays.asList(
                "F01", "F10", "F04", "F20", "F40", "F02", "Q01", "Q45", "Q46", "C06", "C07","K36", "K66", "K92"
        ));

        public static final Set<String> coduriDoarIncasari = new HashSet<>(Arrays.asList(
                "F00", "L40", "L45", "L41", "Q02", "Q15", "Q16", "C09", "K81", "K83", "J50"
        ));

        public static final Set<String> coduriSolduri = new HashSet<>(Arrays.asList(
                "Q50", "Q51", "Q55", "Q56"
        ));

    }
    static HashMap<String, String> sheetReportNameMap = new HashMap<String, String>();
    static HashMap<String, String> sheetRowNameMap = new HashMap<String, String>();
    static HashMap<String, String> sheetColumnNameMap = new HashMap<String, String>();
    static HashMap<String, String> sheetXmlPatternMap = new HashMap<String, String>();

    static HashMap<String, Double> SPD270MapValue1 = new HashMap<String, Double>();
    static HashMap<String, Double> SPD270MapValue2 = new HashMap<String, Double>();

    static HashMap<String, Integer> stringToInt = new HashMap<String, Integer>();
    static HashMap<Integer, String> intToString = new HashMap<Integer, String>();

    static HashSet<String> ignoredColors = new HashSet<String>();

    static List<String> sheetIdList = new LinkedList<String>();
    static List<String> sheetNameList = new LinkedList<String>();
    //SBP
    static List<String> xmlNamesList = new LinkedList<String>();

    static HashMap<String, String> abacusRowNames = new HashMap<String, String>();
    static HashMap<String, String> abacusCountryCodes = new HashMap<String, String>();
    static HashMap<String, String> rpn650XmlMapping = new HashMap<String, String>();
    static HashMap<String, String> rps500XmlMapping = new HashMap<String, String>();


    static boolean getParamsFinrepBnr = true;
    static String senderIDBnr;
    static String sendingDateBnr;
    static String messageTypeBnr;
    static String refDateBnr;
    static String senderMessageIdBnr;
    static String operationTypeBnr;
    static String corelatedWithBnr;
    static String correctionOfBnr;

    static boolean isEncrypted;
    static String encryptKey;
    static String parameterCode;
    static String LEICode;
    static int counter;

    static boolean marker = true;

    /*********** Finrep BNR **************/
    static SimpleDateFormat refDateBnrFormat = new SimpleDateFormat("dd/MM/yyyy");
    static SimpleDateFormat refDateFormat = new SimpleDateFormat("yyyy-MM-dd");

    private static Pattern numericPattern = Pattern.compile("-?\\d+(\\.\\d+)?");

    public static boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        return numericPattern.matcher(strNum).matches();
    }

    public static class flag{
        public static boolean hasItems = false;
    }

    //decripteaza un string folosind encryptKey
    static public String decryptString(String input, String encryptKey) throws Exception
    {
        String ALGO = "AES";
        byte[] keyValue = (String.valueOf(encryptKey) + encryptKey).getBytes();
        Cipher cipher = Cipher.getInstance(ALGO);
        Key key = new SecretKeySpec(keyValue, ALGO);
        cipher.init(2, key);
        byte[] decodedValue = Base64.getDecoder().decode(input);
        byte[] decValue = cipher.doFinal(decodedValue);

        String result = new String(decValue);

        return result;
    }

    //obtine valoarea unei proprietati criptate din config
    static public String getConfigProperty(Properties properties, String propertyName, String encryptKey) throws Exception
    {
        return decryptString(properties.getProperty(propertyName), encryptKey);
    }

    //obtine valoarea unei proprietati necriptate din config
    static public String getConfigProperty(Properties properties, String propertyName) throws Exception
    {
        return properties.getProperty(propertyName);
    }

    private static String getFillColorHex(Cell cell)
    {
        String fillColorString = "";

        if (cell != null)
        {
            CellStyle cellStyle = cell.getCellStyle();
            Color color =  cellStyle.getFillForegroundColorColor();

            if (color instanceof XSSFColor)
            {
                XSSFColor xssfColor = (XSSFColor)color;
                byte[] argb = xssfColor.getARGB();
                fillColorString = Integer.toHexString(((argb[0] & 0xFF) << 24) | ((argb[1] & 0xFF) << 16) | ((argb[2] & 0xFF) << 8) | (argb[3] & 0xFF));

                if (xssfColor.hasTint())
                {
                    byte[] rgb = xssfColor.getRGBWithTint();
                    fillColorString = Integer.toHexString(((argb[0] & 0xFF) << 24) | ((rgb[0] & 0xFF) << 16) | ((rgb[1] & 0xFF) << 8) | (rgb[2] & 0xFF));
                }
            }
            else if (color instanceof HSSFColor)
            {
                HSSFColor hssfColor = (HSSFColor)color;
                short[] rgb = hssfColor.getTriplet();
                fillColorString = Integer.toHexString(0xFF000000 | ((rgb[0] & 0xFF) << 16) | ((rgb[1] & 0xFF) << 8) | (rgb[2] & 0xFF));
            }
        }

        return fillColorString.toUpperCase();
    }


    public static void validateXml(String xsdSchema, String xmlPath)
    {
        if(validationEnabled != null && validationEnabled.equals("Y"))
        {

            StringReader stringReader = new StringReader(xsdSchema);
            StreamSource streamSource = new StreamSource(stringReader);
            String errorMessage = "";

            final List<SAXParseException> exceptions = new LinkedList<SAXParseException>();

            try
            {
                File datafile = new File(xmlPath);

                SchemaFactory factory = SchemaFactory.newInstance(XMLConstants.W3C_XML_SCHEMA_NS_URI);

                Schema schema = factory.newSchema(streamSource);

                Validator validator = schema.newValidator();

                validator.setErrorHandler(new ErrorHandler()
                {
                    @Override
                    public void warning(SAXParseException exception) throws SAXException
                    {
                        exceptions.add(exception);
                    }

                    @Override
                    public void fatalError(SAXParseException exception) throws SAXException
                    {
                        exceptions.add(exception);
                    }

                    @Override
                    public void error(SAXParseException exception) throws SAXException
                    {
                        exceptions.add(exception);
                    }
                });

                validator.validate(new StreamSource(new FileReader(datafile)));

                for(SAXParseException item : exceptions)
                {
                    errorMessage = errorMessage + "Line: " + item.getLineNumber() + " Column: " + item.getColumnNumber() +" Message: " + item.getMessage().replaceAll("cvc-[^ ]+ ", "") + "\n";
                    //System.err.println("Line: " + item.getLineNumber() + " Column: " + item.getColumnNumber() +" Message: " + item.getMessage());
                }

                System.err.println(errorMessage);

                if(stringReader != null)
                    stringReader.close();

                if(reportingEntity.substring(0, 3).equals("302") && !errorMessage.equals(""))
                    System.exit(-1000);
            }
            catch(Exception ex)
            {
                if(stringReader != null)
                    stringReader.close();
            }
        }
    }

    public static boolean isIgnoredColor(Cell cell)
    {
        String hexColorString = getFillColorHex(cell);

		/*
		if(getCellValue(cell).equals("Test Culori"))
			{
				int x = 0;
			}
		*/
        //if(hexColorString != null && !hexColorString.equals(""))
        //System.out.println(hexColorString);

        if (ignoredColors.contains(hexColorString))
            return true;
        else
            return false;
    }

    @SuppressWarnings("deprecation")
    public static String getCellValue(Cell cell)
    {
        if(evaluationFormulasEnabled != null && evaluationFormulasEnabled.equals("Y"))
        {
            try
            {
                evaluator.evaluateFormulaCell(cell);
            }
            catch(Exception ex)
            {

            }
        }

        String cellValue = "";

        if(cell == null || cell.getCellType() == CellType.BLANK || cell.getCellType() == CellType._NONE)
            cellValue = "";
        else if (cell.getCellType() == CellType.NUMERIC)
        {
            //cellValue = dataFormatter.formatRawCellContents(cell.getNumericCellValue(),cell.getCellStyle().getDataFormat(),cell.getCellStyle().getDataFormatString());
            //cellValue = "" + cell.getNumericCellValue();
            //cellValue = decimalFormat.format(cell.getNumericCellValue());
            try
            {
                cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
            }
            catch(Exception ex)
            {
                cellValue = decimalFormat.format(cell.getNumericCellValue());
            }
        }
        else if (cell.getCellType() == CellType.STRING)
            cellValue = cell.getStringCellValue();
        else if (cell.getCellType() == CellType.BOOLEAN) {
            cellValue = Boolean.toString(cell.getBooleanCellValue());
        }
        else if(cell.getCellType() == CellType.FORMULA)
        {
            switch(cell.getCachedFormulaResultType())
            {
                case  NUMERIC :
                    try
                    {
                        cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                    }
                    catch(Exception ex)
                    {
                        cellValue = decimalFormat.format(cell.getNumericCellValue());
                    }
                    break;
                case  STRING :
                    cellValue = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    cellValue = Boolean.toString(cell.getBooleanCellValue());
                    break;
            }
        }

        return cellValue;
    }

    public static String getCellValue(Sheet sheet, String cellAddress)
    {
        try
        {
            CellReference cellRef = new CellReference(cellAddress);
            return getCellValue(sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol()));
        }
        catch(Exception e)
        {
            return "";
        }
    }

    public static String getCellValueFromRow(Row row, int position)
    {
        String val = getCellValue(row.getCell(position));

        return val;
    }


    //pentru extragerea valorii din paranteze ()
    public static String getFormattedCellValue(String cellValue)
    {
        String finalString = cellValue;

        if(cellValue.indexOf("(") >= 0 && cellValue.indexOf("(") < cellValue.indexOf(")"))
            finalString = cellValue.substring(cellValue.indexOf("(") + 1, cellValue.indexOf(")"));
        else
            finalString = cellValue;

        return finalString;
    }


    //verifica daca exista o celula cu valoare pe rand
    public static boolean checkIfRowHasValue(Row row, int startIndex, int endIndex)
    {
        for(int i = startIndex; i <= endIndex; i++)
        {
            if(!getCellValue(row.getCell(i)).trim().equals(""))
                return true;
        }
        return false;
    }


    public static String getStartColumn(String cellRange)
    {
        Matcher matcher = columnPattern.matcher(cellRange.split(":")[0]);
        matcher.find();
        return matcher.group();
    }

    public static String getEndColumn(String cellRange)
    {
        Matcher matcher = columnPattern.matcher(cellRange.split(":")[1]);
        matcher.find();
        return matcher.group();
    }

    public static Integer getStartRow(String cellRange)
    {
        Matcher matcher = rowPattern.matcher(cellRange.split(":")[0]);
        matcher.find();
        return Integer.parseInt(matcher.group()) - 1;
    }

    public static Integer getEndRow(String cellRange)
    {
        Matcher matcher = rowPattern.matcher(cellRange.split(":")[1]);
        matcher.find();
        return Integer.parseInt(matcher.group()) - 1;
    }

    public static void init() throws Exception
    {
        decimalFormat.setMaximumFractionDigits(340);

        //documentName = "Raport " + reportsGroup;
        newLine = System.lineSeparator();

        refDay = Integer.parseInt(refDate.substring(8, 10));
        refMonth = Integer.parseInt(refDate.substring(5, 7));
        refYear = Integer.parseInt(refDate.substring(0, 4));


        //add poi-ooxml*.jar into final jar
        WorkbookFactory.addProvider(new HSSFWorkbookFactory());
        WorkbookFactory.addProvider(new XSSFWorkbookFactory());

        //calea absoluta catre JAR
        path = pathToJar; //new File(XlsToXmlConverter.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getPath();

        //ma asigur ca path se termina cu /
        if(!path.endsWith(File.separator))
            path = path + File.separator;
        //citire credentiale din config.cfg
        if (connectionString.equals("Unknown"))
        {
            //citesc fisierul cu credentialele de conectare la baza de date
            Properties configProperties = new Properties();
            InputStream configInputStream = new FileInputStream(path + "config.cfg");
            configProperties.load(configInputStream);

            if (!isEncrypted)
                encryptKey = null; //config necriptat

            //Connection conn = null;

            //incerc pe rand sa ma conectez la toate bazele de date mentionate in config, pana reusesc conectarea la una din ele
            boolean isConnected = false;
            int propertyNum = 1;

            while (isConnected == false)
            {
                try
                {
                    String propertyNumString;

                    //pentru primul set de proprietati nu mai punem sufix 1
                    if(propertyNum == 1)
                        propertyNumString = "";
                    else
                        propertyNumString = String.valueOf(propertyNum);

                    if (encryptKey != null && !encryptKey.equals("QUwSE9rc")) //config criptat, nu cu cheia de la alpha
                    {
                        JDBC_DRIVER = getConfigProperty(configProperties, "A" + propertyNumString, encryptKey);
                        DB_URL = getConfigProperty(configProperties, "B" + propertyNumString, encryptKey);
                        USER = getConfigProperty(configProperties, "C" + propertyNumString, encryptKey);
                        PASS = getConfigProperty(configProperties, "D" + propertyNumString, encryptKey);
                    }

                    else //config necriptat
                    {
                        JDBC_DRIVER = getConfigProperty(configProperties, "JDBC_DRIVER" + propertyNumString);
                        DB_URL = getConfigProperty(configProperties, "DB_URL" + propertyNumString);
                        USER = getConfigProperty(configProperties, "USER" + propertyNumString);
                        PASS = getConfigProperty(configProperties, "PASS" + propertyNumString);
                    }

                    //incerc sa ma conectez la baza de date
                    Class.forName(JDBC_DRIVER);

                    if (DB_URL.contains("integratedSecurity=true")) {
                        String currentUser = System.getProperty("user.name");
                        connection = DriverManager.getConnection(DB_URL);
                    } else {
                        connection = DriverManager.getConnection(DB_URL, USER, PASS);
                    }
                    //connection = DriverManager.getConnection(DB_URL, USER, PASS);

                    isConnected = true;
                    System.err.println("Database connection established successfully: " + DB_URL);
                }
                catch(Exception e)
                {
                    //daca sunt toate null atunci am epuizat toate bazele de date din config
                    if(JDBC_DRIVER == null && DB_URL == null && USER == null && PASS == null)
                        throw e;
                    System.err.println("Database connection failed: " + e.getMessage());
                    e.printStackTrace();

                    propertyNum++;
                }
            }
			/*
			if(isEncrypted == false)
			{
				JDBC_DRIVER = configProperties.getProperty("JDBC_DRIVER");
				DB_URL = configProperties.getProperty("DB_URL");
				USER = configProperties.getProperty("USER");
				PASS = configProperties.getProperty("PASS");
			}
			else if(isEncrypted == true && encryptKey != null && !encryptKey.equals(""))
			{
				JDBC_DRIVER = configProperties.getProperty("A");
				DB_URL = configProperties.getProperty("B");
				USER = configProperties.getProperty("C");
				PASS = configProperties.getProperty("D");


				try {
					String ALGO = "AES";
					byte[] keyValue = (String.valueOf(encryptKey) + encryptKey).getBytes();
			        Cipher x = Cipher.getInstance(ALGO);
			        Key key = new SecretKeySpec(keyValue, ALGO);
			        x.init(2, key);
			        byte[] decordedValue = Base64.getDecoder().decode(JDBC_DRIVER);
			        byte[] decValue = x.doFinal(decordedValue);
			        JDBC_DRIVER = new String(decValue);
			        decordedValue =  Base64.getDecoder().decode(DB_URL);
			        decValue = x.doFinal(decordedValue);
			        DB_URL = new String(decValue);
			        decordedValue =  Base64.getDecoder().decode(USER);
			        decValue = x.doFinal(decordedValue);
			        USER = new String(decValue);
			        decordedValue =  Base64.getDecoder().decode(PASS);
			        decValue = x.doFinal(decordedValue);
			        PASS = new String(decValue);
				}
				catch(Exception e)
				{
					e.printStackTrace(new PrintWriter(sw));
				    System.err.println("Exception at decrypt with key: " + encryptKey + " grup: " + reportsGroup + " data: " + + " path: " + pathToJar + " nume: " + xlsFileName + " \n" + sw.toString());

				}

			}
			else
			{
				JDBC_DRIVER = configProperties.getProperty("JDBC_DRIVER");
				DB_URL = configProperties.getProperty("DB_URL");
				USER = configProperties.getProperty("USER");
				PASS = configProperties.getProperty("PASS");
			}

			*/

            configInputStream.close();
        }
        else
        {
            String credentials = decryptString(connectionString, encryptKey);

            String regex = "JDBC_DRIVER=(.+);DB_URL=(.+);USER=(.+);PASS=(.+)";

            Pattern pattern = Pattern.compile(regex);

            Matcher matcher = pattern.matcher(credentials);

            if (matcher.find()) {
                JDBC_DRIVER = matcher.group(1);
                DB_URL = matcher.group(2);
                USER = matcher.group(3);
                PASS = matcher.group(4);

                //ma conectez la baza de date
                Class.forName(JDBC_DRIVER);
                connection = DriverManager.getConnection(DB_URL, USER, PASS);
            }

            else System.out.println("Couldn't resolve connection string.");


        }
        statement = connection.createStatement();
        statement_seq = connection.createStatement();
        sqlQuery_seq = "select next value for senderMessage_s";

        SenderMessageId = statement_seq.executeQuery(sqlQuery_seq);
        if (SenderMessageId.next()) {
            SenderMSGID = SenderMessageId.getString(1);
        }


        if(DB_URL.contains("sqlserver"))
            refDateSql = "convert(DATE,'" + refDate + "', 126)";
        else
            refDateSql = "to_date('" + refDate + "', 'yyyy-mm-dd')";

        //obtin id-ul si numele fisierului excel asignat grupului de rapoarte, valid la data de referinta
		/*sqlQuery = "select xf.file_id," + newLine +
				   "       xf.file_ignored_colors" + newLine +
           		   "  from rep_reports_groups rg," + newLine +
           		   "       rep_xls_grp_mapping xm," + newLine +
           		   "       rep_xls_files xf" + newLine +
           		   " where rg.reports_group_code = '" + reportsGroup + "'" + newLine +
           		   "   and xf.file_name = '" + xlsFileName + "'" + newLine +
           		   "   and rg.reports_group_id = xm.reports_group_id" + newLine +
           		   "   and " + refDateSql + " between xm.valid_from and xm.valid_to" + newLine +
           		   "   and xf.file_id = xm.file_id" + newLine +
           		   "   and " + refDateSql + " between xf.valid_from and xf.valid_to";*/
        sqlQuery = "select xf.file_id," + newLine +
                "       xf.file_ignored_colors," + newLine +
                "       xf.document_name," + newLine +
                "       xf.validation_enabled," + newLine +
                "       xf.xsd_schema," + newLine +
                "       xf.evaluate_functions" + newLine +
                "  from rep_xls_files xf" + newLine +
                " where xf.file_name = '" + xlsFileName + "'" + newLine +
                "   and " + refDateSql + " between xf.valid_from and xf.valid_to";

        resultSet = statement.executeQuery(sqlQuery);

        resultSet.next();

        xlsFileId = resultSet.getInt(1) + "";
        xlsFileIgnoredColors = resultSet.getString(2);
        documentName = resultSet.getString(3);
        validationEnabled = resultSet.getString(4);
        evaluationFormulasEnabled = resultSet.getString(6);

        if(DB_URL.contains("sqlserver"))
        {
            xsdSchema = resultSet.getString(5);
        }
        else
        {
            clob = resultSet.getClob(5);

            if(clob != null)
                xsdSchema = clob.getSubString(1, (int)clob.length());
        }

        if(xlsFileIgnoredColors != null)
            for(String color : xlsFileIgnoredColors.split(","))
                ignoredColors.add(color);

        if(xlsFileId.equals("20"))
        {
            sqlQuery = "select rcd.clasification_det_code,"+ newLine +
                    "rcd.attr1,"+ newLine +
                    "rcd.attr2,"+ newLine +
                    "rcd.attr3" + newLine +
                    " from rep_clasifications rc" + newLine +
                    " join rep_clasifications_det rcd on rc.clasification_id = rcd.clasification_id" + newLine +
                    "   where rc.clasification_code = 'SPD270_CONTACT_DETAILS'";

            resultSetContacts = statement.executeQuery(sqlQuery);
            while(resultSetContacts.next())
            {
                SPD270_Contacts[nrContacte][0] = resultSetContacts.getString(1);
                SPD270_Contacts[nrContacte][1] = resultSetContacts.getString(2);
                SPD270_Contacts[nrContacte][2] = resultSetContacts.getString(3);
                SPD270_Contacts[nrContacte][3] = resultSetContacts.getString(4);

                nrContacte++;
            }
        }
        if(xlsFileId.equals("21"))
        {
            sqlQuery = "select rcd.clasification_det_code,"+ newLine +
                    "rcd.attr1,"+ newLine +
                    "rcd.attr2,"+ newLine +
                    "rcd.attr3" + newLine +
                    " from rep_clasifications rc" + newLine +
                    " join rep_clasifications_det rcd on rc.clasification_id = rcd.clasification_id" + newLine +
                    "   where rc.clasification_code = 'SPD279_CONTACT_DETAILS'";

            resultSetContacts = statement.executeQuery(sqlQuery);
            while(resultSetContacts.next())
            {
                SPD279_Contacts[nrContacte][0] = resultSetContacts.getString(1);
                SPD279_Contacts[nrContacte][1] = resultSetContacts.getString(2);
                SPD279_Contacts[nrContacte][2] = resultSetContacts.getString(3);
                SPD279_Contacts[nrContacte][3] = resultSetContacts.getString(4);

                nrContacte++;
            }
        }



        //obtin id-urile si numele sheet-urilor din fisierul excel
        sqlQuery = "select xs.sheet_id," + newLine +
                "       xs.sheet_name," + newLine +
                "       xs.sheet_report_name," + newLine +
                "       xs.sheet_row_name," + newLine +
                "       xs.sheet_column_name," + newLine +
                "       xs.sheet_xml_pattern," + newLine +
                "       xs.sheet_report_scope," + newLine +
                "       xs.sheet_report_frequency" + newLine +
                "  from rep_xls_sheets xs" + newLine +
                " where xs.file_id = " + xlsFileId + newLine +
                "   and " + refDateSql + " between xs.valid_from and xs.valid_to" + newLine +
                " order by xs.sheet_order_no";

        resultSet = statement.executeQuery(sqlQuery);






        //adaug id-urile si numele sheet-urilor in 2 liste
        while(resultSet.next())
        {
            String sheetId = resultSet.getString(1);
            String sheetName = resultSet.getString(2);
            String sheetReportName = resultSet.getString(3);
            String sheetRowName = resultSet.getString(4);
            String sheetColumnName = resultSet.getString(5);
            String sheetXmlPattern = resultSet.getString(6);
            String sheetReportScope = resultSet.getString(7);
            String sheetReportFrequency = resultSet.getString(8);


            //sar sheet-urile raportate doar in consolidat
            if (consolidatedFlag.equals("N") && sheetReportScope.equals("g"))
                continue;

            //sar sheet-urile care nu se raporteaza la data de referinta
            if (sheetReportFrequency.equals("a") && refMonth != 12)
                continue;
            if (sheetReportFrequency.equals("s") && refMonth != 6 && refMonth != 12)
                continue;
            if (sheetReportFrequency.equals("q") && refMonth != 3 && refMonth != 6 && refMonth != 9 && refMonth != 12)
                continue;


            sheetIdList.add(sheetId);
            sheetNameList.add(sheetName);

            sheetReportNameMap.put(/*sheetName*/sheetId, sheetReportName);
            sheetRowNameMap.put(/*sheetName*/sheetId, sheetRowName);
            sheetColumnNameMap.put(/*sheetName*/sheetId, sheetColumnName);
            sheetXmlPatternMap.put(/*sheetName*/sheetId, sheetXmlPattern);
        }

        //iau codul bancii din parametrii generali
        if(reportsGroup != null && reportsGroup.equals("#aed#"))
            parameterCode = "REPORTING_ENTITY_RAIF";
        else
            parameterCode = "REPORTING_ENTITY";

        sqlQuery = "select t.parameter_value" + newLine +
                "  from rep_general_parameters t" + newLine +
                " where t.parameter_code = '" + parameterCode + "'";

        resultSet = statement.executeQuery(sqlQuery);
        resultSet.next();
        reportingEntity = resultSet.getString(1);

        //iau codul LEI
        sqlQuery = "select t.parameter_value" + newLine +
                "  from rep_general_parameters t" + newLine +
                " where t.parameter_code = '" + "LEICODE" + "'";

        resultSet = statement.executeQuery(sqlQuery);
        resultSet.next();
        LEICode = resultSet.getString(1);


        if(xlsFileId.equals("17"))
            reportingEntity = reportingEntity + "";
        else if (consolidatedFlag.equals("Y") == false)
        {
            if( (xlsFileId.equals("6") ||
                    xlsFileId.equals("7") ||
                    xlsFileId.equals("9")) && !xlsFileName.contains("CCY") /*||
			   xlsFileId.equals("8") ||
			   xlsFileId.equals("2") ||
			   xlsFileName.equals("Annex 8 (Large exposures).xls") || xlsFileId.equals("4")*/)
                reportingEntity = reportingEntity + "r";
            else
                reportingEntity = reportingEntity + "i";
        }
        else if(consolidatedFlag.equals("Y") == true && ((xlsFileId.equals("6") || xlsFileId.equals("14") ||
                xlsFileId.equals("7") ||
                xlsFileId.equals("9")) && !xlsFileName.contains("CCY")  /*||
														 xlsFileId.equals("2") || xlsFileId.equals("8") ||
														 xlsFileName.equals("Annex 8 (Large exposures).xls") || xlsFileId.equals("4")*/))
            reportingEntity = reportingEntity + "q";
        else if(consolidatedFlag.equals("Y") == true && reportingEntity.equals("317") == true && xlsFileId.equals("8"))
            reportingEntity = reportingEntity + "i";


        int nrCrt = 0;

        for(char c1 = 'A'; c1 <= 'Z'; c1++)
        {
            stringToInt.put("" + c1, nrCrt);
            intToString.put(nrCrt, "" + c1);
            nrCrt++;
        }

        for(char c1 = 'A'; c1 <= 'Z'; c1++)
        {
            for(char c2 = 'A'; c2 <= 'Z'; c2++)
            {
                stringToInt.put("" + c1 + c2, nrCrt);
                intToString.put(nrCrt, "" + c1 + c2);
                nrCrt++;
            }
        }

        for(char c1 = 'A'; c1 <= 'Z'; c1++)
        {
            for(char c2 = 'A'; c2 <= 'Z'; c2++)
            {
                for(char c3 = 'A'; c3 <= 'Z'; c3++)
                {
                    stringToInt.put("" + c1 + c2 + c3, nrCrt);
                    intToString.put(nrCrt, "" + c1 + c2 + c3);
                    nrCrt++;
                }
            }
        }

        //deschid fisierul excel
        workbook = WorkbookFactory.create(new FileInputStream(path + xlsFileName));


        evaluator = workbook.getCreationHelper().createFormulaEvaluator();


        if(xlsFileId.equals("11")  && xlsFileName.startsWith("MACHETA")) //Xml Bnr
        {
            //xmlFileName = "RFC400-05.xml";
            xmlFileName = xlsFileName.substring(0, xlsFileName.lastIndexOf(".") + 1) + "xml";
        }
        else if(xlsFileId.equals("16"))
        {

            xmlFileName = refDate.replace("-", "") + "_" + reportingEntity + "_SBP_CR_ALL_ITS.xml";
            xmlFileNameIMV = refDate.replace("-", "") + "_" + reportingEntity + "_SBP_IMV_ALL_ITS.xml";
            xmlFileNameRM = refDate.replace("-", "") + "_" + reportingEntity + "_SBP_RM_ALL_ITS.xml";

            xmlNamesList.add(path + xmlFileName);
            xmlNamesList.add(path + xmlFileNameIMV);
            xmlNamesList.add(path + xmlFileNameRM);
        }
        else if (xlsFileId.equals("58") &&  hasErrors)
            xmlFileName = "ERORI_" + xlsFileName.substring(0, xlsFileName.lastIndexOf(".") + 1).replace("MACHETA_","") + "xml";
        else if (xlsFileId.equals("58") && !hasErrors)
            xmlFileName = xlsFileName.substring(0, xlsFileName.lastIndexOf(".") + 1).replace("MACHETA_","") + "xml";
        else
            xmlFileName = xlsFileName.substring(0, xlsFileName.lastIndexOf(".") + 1) + "xml";

        //deschid fisierul xml de output
        xmlWriter = new PrintWriter(path + xmlFileName, "UTF-8");

        //Adaugat pentru generare SBP
        if(xlsFileId.equals("16"))
        {
            xmlWriterIMV = new PrintWriter(path + xmlFileNameIMV, "UTF-8");
            xmlWriterRM = new PrintWriter(path + xmlFileNameRM, "UTF-8");
        }

    }

    public static void writeXml(String line)
    {
        xmlWriter.println(line.replace("&", "&amp;"));
    }

    public static void writeXmlIMV(String line)
    {
        xmlWriterIMV.println(line.replace("&", "&amp;"));
    }

    public static void writeXmlRM(String line)
    {
        xmlWriterRM.println(line.replace("&", "&amp;"));
    }

    private static void zipFiles(List<String> xmlFiles, String fileName) throws IOException
    {

        String zipFileName = fileName.replace(".xlsx", ".zip");///firstFile.getName().concat(".zip");

        FileOutputStream fos = new FileOutputStream(zipFileName);
        ZipOutputStream zos = new ZipOutputStream(fos);

        for (int i = 0; i < xmlFiles.size(); i++)
        {

            zos.putNextEntry(new ZipEntry(new File(xmlFiles.get(i)).getName()));

            byte[] bytes = Files.readAllBytes(Paths.get(xmlFiles.get(i)));
            zos.write(bytes, 0, bytes.length);
            zos.closeEntry();
        }

        zos.close();
    }

    private static String dateFromNumber(String number)
    {
        return  new SimpleDateFormat("yyyy-MM-dd").format(DateUtil.getJavaDate(Double.parseDouble(number)));
    }

    //cazul COREP - 6.2 - Group Solvency - sheet dinamic & rand xml = rand macheta
    public static void writeXmlCorep6_2()
    {
        int currentDynamicRow = 1;

        while(getCellValue(sheet, startColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String currentXmlTag = xmlPattern;
            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode;

                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                currentXmlTag = currentXmlTag.replace("@F011", String.format("%03d", currentDynamicRow));

                if (currentXmlTag != null)
                {
                    if(colCode.equals("010"))
                        currentXmlTag = currentXmlTag.replace("@ENTITY_FULLNAME", cellValue);
                    else if(colCode.equals("020"))
                        currentXmlTag = currentXmlTag.replace("@GROUP_CODE", cellValue);
                    else if(colCode.equals("025"))
                        currentXmlTag = currentXmlTag.replace("@LEI_CODE", cellValue);
                    else if(colCode.equals("030"))
                        currentXmlTag = currentXmlTag.replace("@INSTITUTION_OR_EQUIVALENT", cellValue);
                    else if(colCode.equals("035"))
                        currentXmlTag = currentXmlTag.replace("@SCOPE_OF_DATA", cellValue);
                    else if(colCode.equals("040"))
                        currentXmlTag = currentXmlTag.replace("@Iso_code", cellValue);
                    else if(colCode.equals("050"))
                        currentXmlTag = currentXmlTag.replace("@SHARE_OF_HOLDING", cellValue);
                    else if(colCode.equals("060"))
                        currentXmlTag = currentXmlTag.replace("@RISK_EXPOSURE_TOTAL", cellValue);
                    else if(colCode.equals("070"))
                        currentXmlTag = currentXmlTag.replace("@CREDIT_RISK", cellValue);
                    else if(colCode.equals("080"))
                        currentXmlTag = currentXmlTag.replace("@FX_COMMODITIES_RISK", cellValue);
                    else if(colCode.equals("090"))
                        currentXmlTag = currentXmlTag.replace("@OPERATIONAL_RISK", cellValue);
                    else if(colCode.equals("100"))
                        currentXmlTag = currentXmlTag.replace("@OTHER_RISK", cellValue);
                    else if(colCode.equals("110"))
                        currentXmlTag = currentXmlTag.replace("@own_funds", cellValue);
                    else if(colCode.equals("120"))
                        currentXmlTag = currentXmlTag.replace("@OWN_FUNDS_QUALIFYING", cellValue);
                    else if(colCode.equals("130"))
                        currentXmlTag = currentXmlTag.replace("@OWN_FUNDS_INSTRUMENT", cellValue);
                    else if(colCode.equals("140"))
                        currentXmlTag = currentXmlTag.replace("@T1_CAPITAL", cellValue);
                    else if(colCode.equals("150"))
                        currentXmlTag = currentXmlTag.replace("@T1_QUALIFYING", cellValue);
                    else if(colCode.equals("160"))
                        currentXmlTag = currentXmlTag.replace("@T1_OWN_FUNDS_INSTRUMENTS", cellValue);
                    else if(colCode.equals("170"))
                        currentXmlTag = currentXmlTag.replace("@CET1_CAPITAL", cellValue);
                    else if(colCode.equals("180"))
                        currentXmlTag = currentXmlTag.replace("@CET1_MINORITY_INTERESTS", cellValue);
                    else if(colCode.equals("190"))
                        currentXmlTag = currentXmlTag.replace("@CET1_OWN_FUNDS_INSTRUMENTS", cellValue);
                    else if(colCode.equals("200"))
                        currentXmlTag = currentXmlTag.replace("@AT1_CAPITAL", cellValue);
                    else if(colCode.equals("210"))
                        currentXmlTag = currentXmlTag.replace("@AT1_QUALIFYING", cellValue);
                    else if(colCode.equals("220"))
                        currentXmlTag = currentXmlTag.replace("@T2_CAPITAL", cellValue);
                    else if(colCode.equals("230"))
                        currentXmlTag = currentXmlTag.replace("@T2_QUALIFYING", cellValue);
                    else if(colCode.equals("240"))
                        currentXmlTag = currentXmlTag.replace("@RISK_EXPOSURE_AMOUNT_TOTAL", cellValue);
                    else if(colCode.equals("250"))
                        currentXmlTag = currentXmlTag.replace("@RISK_EXPOSURE_CREDIT_RISK", cellValue);
                    else if(colCode.equals("260"))
                        currentXmlTag = currentXmlTag.replace("@RISK_EXPOSURE_FX_COMMODITIES", cellValue);
                    else if(colCode.equals("270"))
                        currentXmlTag = currentXmlTag.replace("@RISK_EXPOSURE_OP", cellValue);
                    else if(colCode.equals("280"))
                        currentXmlTag = currentXmlTag.replace("@RISK_EXPOSURE_OTHER", cellValue);
                    else if(colCode.equals("290"))
                        currentXmlTag = currentXmlTag.replace("@QUALIFYING_OWN_FUNDS", cellValue);
                    else if(colCode.equals("300"))
                        currentXmlTag = currentXmlTag.replace("@qualifying_t1", cellValue);
                    else if(colCode.equals("310"))
                        currentXmlTag = currentXmlTag.replace("@QUALIFYING_MINORITY_INTERESTS", cellValue);
                    else if(colCode.equals("320"))
                        currentXmlTag = currentXmlTag.replace("@QUALIFYING_T1_ADDITIONAL", cellValue);
                    else if(colCode.equals("330"))
                        currentXmlTag = currentXmlTag.replace("@QUALIFYING_T2", cellValue);
                    else if(colCode.equals("340"))
                        currentXmlTag = currentXmlTag.replace("@QUALIFYING_GOODWILL", cellValue);
                    else if(colCode.equals("350"))
                        currentXmlTag = currentXmlTag.replace("@CONSOLIDATED_OWN_FUNDS", cellValue);
                    else if(colCode.equals("360"))
                        currentXmlTag = currentXmlTag.replace("@CONSOLIDATED_COMMPN_EQUITY_T1", cellValue);
                    else if(colCode.equals("370"))
                        currentXmlTag = currentXmlTag.replace("@CONSOLIDATED_ADDITIONAL_T1", cellValue);
                    else if(colCode.equals("380"))
                        currentXmlTag = currentXmlTag.replace("@CONSOLIDATED_CONRIBUT_CONSOL", cellValue);
                    else if(colCode.equals("390"))
                        currentXmlTag = currentXmlTag.replace("@CONSOLIDATED_GOODWILL", cellValue);
                    else if(colCode.equals("400"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_REQUIREMENT", cellValue);
                    else if(colCode.equals("410"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_CONSERVATION", cellValue);
                    else if(colCode.equals("420"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_INST_SPEC_CC", cellValue);
                    else if(colCode.equals("430"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_MEMBER_STATE", cellValue);
                    else if(colCode.equals("440"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_SYSTEMICAL_RISK", cellValue);
                    else if(colCode.equals("450"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_SYSTEMICAL_IMP_INST", cellValue);
                    else if(colCode.equals("470"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_GLOBAL_IMP_INST", cellValue);
                    else if(colCode.equals("480"))
                        currentXmlTag = currentXmlTag.replace("@CAPITAL_OTHER_IMP_INST", cellValue);
                }
            }

            writeXml(currentXmlTag);

            currentDynamicRow ++;
        }
    }

    //cazul COREP - 8.2 - CR IRB - Obligor Pools - dinamic (nr variabil de randuri)
    public static void writeXmlCorep8_2()
    {
        int currentDynamicRow = 1;

        while(getCellValue(sheet, startColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String rowCode, xmlRowCode = "";

            xmlRowCode = xmlRowNamePattern;

            //String currentXmlTagDummyDynCell = xmlPattern;
            //currentXmlTagDummyDynCell = currentXmlTagDummyDynCell.replace("@f001", reportingEntity);
            //currentXmlTagDummyDynCell =

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {
                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {
                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                    currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                    //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                    currentXmlTag = currentXmlTag.replace("@I056", valueOfI056);

                    currentXmlTag = currentXmlTag.replace("@B790", "" + currentDynamicRow);


                    if(currentXmlTag.indexOf("allocated_regulator_text") != -1)
                        currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                    else
                        currentXmlTag = currentXmlTag.replace("@TEXT", "#TEXT");


                    if((xlsFileId.equals("1") || xlsFileName.equals("Annex_1_(Solvency).xlsx")) && currentSheetName.startsWith("8"))
                    {
                        if(tableB015Value != null)
                            currentXmlTag = currentXmlTag.replace("@B015", tableB015Value);
                        else
                            currentXmlTag = currentXmlTag.replace(" B015=\"@B015\"", "");

                        if(currentSheetName.startsWith("8.1") && xmlRowCode.equals("015"))
                            currentXmlTag = currentXmlTag.replace("@B046", "1");
                        else
                            currentXmlTag = currentXmlTag.replace(" B046=\"@B046\"", "");
                    }


                    writeXml(currentXmlTag);
                }
            }

            currentDynamicRow++;
        }
    }

    //cazul COREP - 14 - Securitisation Details - sheet dinamic & rand xml = rand macheta
    public static void writeXmlCorep14()
    {
        int currentDynamicRow = 1;
        String sec502 = "";

        if(currentSheetName.equals("14.1"))
            sec502 = valueOfB272;

        while(getCellValue(sheet, startColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String currentXmlTag = xmlPattern;
            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode;

                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                if (currentXmlTag != null)
                {
                    if(currentSheetName.equals("14"))
                    {
                        if(colCode.equals("005"))
                            currentXmlTag = currentXmlTag; //nu se pune nicaieri???
                        else if(colCode.equals("010"))
                            currentXmlTag = currentXmlTag.replace("@CODE", cellValue);
                        else if(colCode.equals("020"))
                            currentXmlTag = currentXmlTag.replace("@SECURITISATION_ID", cellValue);
                        else if(colCode.equals("021"))
                            currentXmlTag = currentXmlTag.replace("@SECURITISATION_PLACEMENT", cellValue);
                        else if(colCode.equals("110"))
                            currentXmlTag = currentXmlTag.replace("@INSTITUTION_ROLE", cellValue);
                        else if(colCode.equals("030"))
                            currentXmlTag = currentXmlTag.replace("@ORIGINATOR_ID", cellValue);
                        else if(colCode.equals("040"))
                            currentXmlTag = currentXmlTag.replace("@SECURITISATION_TYPE", cellValue);
                        else if(colCode.equals("051"))
                            currentXmlTag = currentXmlTag.replace("@ACCOUNTING_TREATMENT", cellValue);
                        else if(colCode.equals("060"))
                            currentXmlTag = currentXmlTag.replace("@SOLVENCY_TREATMENT", cellValue);
                        else if(colCode.equals("061"))
                            currentXmlTag = currentXmlTag.replace("@SIGN_RISK_TRANSFER", cellValue);
                        else if(colCode.equals("070"))
                            currentXmlTag = currentXmlTag.replace("@RE_SECURITISATION", cellValue);
                        else if(colCode.equals("075"))
                            currentXmlTag = currentXmlTag.replace("@STS_SECURITISATION", cellValue);
                        else if(colCode.equals("446"))
                            currentXmlTag = currentXmlTag.replace("@STS_SQ_DIFF_CAPITAL_TREATMENT", cellValue);
                        else if(colCode.equals("080"))
                            currentXmlTag = currentXmlTag.replace("@RETENTION_TYPE", cellValue);
                        else if(colCode.equals("090"))
                            currentXmlTag = currentXmlTag.replace("@RETENTION_PERCENT", cellValue);
                        else if(colCode.equals("100"))
                            currentXmlTag = currentXmlTag.replace("@RETENTION_COMPLIANCE", cellValue);
                        else if(colCode.equals("120"))
                            currentXmlTag = currentXmlTag.replace("@NON_ABCP_ORIGINATION_DATE", cellValue);
                        else if(colCode.equals("121"))
                            currentXmlTag = currentXmlTag.replace("@NON_ABCP_LAST_ISSUANCE_DATE", cellValue);
                        else if(colCode.equals("130"))
                            currentXmlTag = currentXmlTag.replace("@NON_ABCP_ORIGINATION_AMOUNT", cellValue);
                        else if(colCode.equals("140"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_TOTAL", cellValue);
                        else if(colCode.equals("150"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_INSTITUTIONS_SHARE", cellValue);
                        else if(colCode.equals("160"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_TYPE", cellValue);
                        else if(colCode.equals("171"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_APPROACH_APPLIED", cellValue);
                        else if(colCode.equals("180"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_NUMBER_CODE", cellValue);
                        else if(colCode.equals("181"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_DEFAULT_W", cellValue);
                        else if(colCode.equals("190"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_COUNTRY", cellValue);
                        else if(colCode.equals("201"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_ELGD", cellValue);
                        else if(colCode.equals("202"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_EL", cellValue);
                        else if(colCode.equals("203"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_UL", cellValue);
                        else if(colCode.equals("204"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_AVG_MATURITY_ASSETS", cellValue);
                        else if(colCode.equals("210"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_ADJUSTMENTS_PROVISIONS", cellValue);
                        else if(colCode.equals("221"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_OWF_REQ_BEFORE_KIRB", cellValue);
                        else if(colCode.equals("222"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_RETAIL_IN_IRB_POOL", cellValue);
                        else if(colCode.equals("223"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_OWF_REQ_BEFORE_KSA", cellValue);
                        else if(colCode.equals("225"))
                            currentXmlTag = currentXmlTag.replace("@SEC_EXP_ADJUSTMENTS_CREDITRISK", cellValue);
                        else if(colCode.equals("230"))
                            currentXmlTag = currentXmlTag.replaceFirst("@SEC_STR_ON_SENIOR", cellValue);
                        else if(colCode.equals("231"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_ON_SENIOR_ATTACHMENT_P", cellValue);
                        else if(colCode.equals("232"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_ON_SENIOR_CQS", cellValue);
                        else if(colCode.equals("240"))
                            currentXmlTag = currentXmlTag.replaceFirst("@SEC_STR_ON_MEZZANINE", cellValue);
                        else if(colCode.equals("241"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_ON_MEZZANINE_TRANCHES", cellValue);
                        else if(colCode.equals("242"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_ON_MEZZANINE_CQS", cellValue);
                        else if(colCode.equals("250"))
                            currentXmlTag = currentXmlTag.replaceFirst("@SEC_STR_ON_FIRST_LOSS", cellValue);
                        else if(colCode.equals("251"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_ON_FIRST_LOSS_DETACHM", cellValue);
                        else if(colCode.equals("252"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_ON_FIRST_LOSS_CQS", cellValue);
                        else if(colCode.equals("260"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_OFF_SENIOR", cellValue);
                        else if(colCode.equals("270"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_OFF_MEZZANINE", cellValue);
                        else if(colCode.equals("280"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_OFF_FIRST_LOSS", cellValue);
                        else if(colCode.equals("290"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_TERMINATION_DATE", cellValue);
                        else if(colCode.equals("291"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_MATURITY_ORIG_CALL_OPT", cellValue);
                        else if(colCode.equals("300"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_FINAL_MATURITY_DATE", cellValue);
                        else if(colCode.equals("302"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_MEM_ATTACHM_RISK_SOLD", cellValue);
                        else if(colCode.equals("303"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_MEM_DETACHM_RISK_SOLD", cellValue);
                        else if(colCode.equals("304"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR_MEM_RISK_TRANSFER", cellValue);


                    }
                    else if(currentSheetName.equals("14.1"))
                    {
                        if(colCode.equals("005"))
                            currentXmlTag = currentXmlTag; //nu se pune nicaieri???
                        else if(colCode.equals("010"))
                            currentXmlTag = currentXmlTag.replace("@CODE", cellValue);
                        else if(colCode.equals("020"))
                            currentXmlTag = currentXmlTag.replace("@SECURITISATION_ID", cellValue);
                        else if(colCode.equals("310"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_ON_SENIOR", cellValue);
                        else if(colCode.equals("320"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_ON_MEZZANINE", cellValue);
                        else if(colCode.equals("330"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_ON_FIRST_LOSS", cellValue);
                        else if(colCode.equals("340"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_OFF_SENIOR", cellValue);
                        else if(colCode.equals("350"))
                            currentXmlTag = currentXmlTag.replaceFirst("@SEC_STR2_OFF_MEZZANINE", cellValue);
                        else if(colCode.equals("351"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_OFF_MEZZANINE_RW", cellValue);
                        else if(colCode.equals("360"))
                            currentXmlTag = currentXmlTag.replaceFirst("@SEC_STR2_OFF_FIRST_LOSS", cellValue);
                        else if(colCode.equals("361"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_OFF_FIRST_LOSS_RW", cellValue);
                        else if(colCode.equals("370"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_CREDIT_SUBSTITUTES", cellValue);
                        else if(colCode.equals("380"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_IRS_CRS", cellValue);
                        else if(colCode.equals("390"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_LIQUIDITY_FACILITIES", cellValue);
                        else if(colCode.equals("400"))
                            currentXmlTag = currentXmlTag.replace("@SEC_STR2_OTHER", cellValue);
                        else if(colCode.equals("411"))
                            currentXmlTag = currentXmlTag.replace("@EXPOSURE_VALUE", cellValue);
                        else if(colCode.equals("420"))
                            currentXmlTag = currentXmlTag.replace("@EXP_VALUE_DEDUCTED_OWN_FUNDS", cellValue);
                        else if(colCode.equals("430"))
                            currentXmlTag = currentXmlTag.replace("@TOTAL_OWN_BEFORE_CAP", cellValue);
                        else if(colCode.equals("431"))
                            currentXmlTag = currentXmlTag.replace("@TOTAL_OWN_RISK_WEIGHT_CAP", cellValue);
                        else if(colCode.equals("432"))
                            currentXmlTag = currentXmlTag.replace("@TOTAL_OWN_OVERALL_CAP", cellValue);
                        else if(colCode.equals("440"))
                            currentXmlTag = currentXmlTag.replace("@TOTAL_OWN_AFTER_CAP", cellValue);
                        else if(colCode.equals("447"))
                            currentXmlTag = currentXmlTag.replace("@MEM_RWA_ERBA", cellValue);
                        else if(colCode.equals("448"))
                            currentXmlTag = currentXmlTag.replace("@MEM_RWA_SA", cellValue);
                        else if(colCode.equals("450"))
                            currentXmlTag = currentXmlTag.replace("@TRADING_BOOK_NON_CTP", cellValue);
                        else if(colCode.equals("460"))
                            currentXmlTag = currentXmlTag.replace("@TRADING_BOOK_NETPOS_LONG", cellValue);
                        else if(colCode.equals("470"))
                            currentXmlTag = currentXmlTag.replace("@TRADING_BOOK_NETPOS_SHORT", cellValue);


                    }
                }
            }

            currentXmlTag = currentXmlTag.replace("@SEC502", sec502);
            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);

            currentXmlTag = currentXmlTag.replaceAll(" [^ ]+@[^ ]+", "");
            writeXml(currentXmlTag);

            currentDynamicRow++;
        }
    }

    //cazul COREP - 10.2 - CR EQU IRB - Obligor Pools - dinamic (nr variabil de randuri)
    public static void writeXmlCorep10_2()
    {
        int currentDynamicRow = 1;

        while(getCellValue(sheet, startColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String rowCode, xmlRowCode = "";

            xmlRowCode = xmlRowNamePattern;

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {
                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {
                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                    currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                    //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                    currentXmlTag = currentXmlTag.replace("@I056", valueOfI056);

                    currentXmlTag = currentXmlTag.replace("@B790", "" + currentDynamicRow);


                    if(currentXmlTag.indexOf("allocated_regulator_text") != -1)
                        currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                    else
                        currentXmlTag = currentXmlTag.replace("@TEXT", "#TEXT");

                    writeXml(currentXmlTag);
                }
            }

            currentDynamicRow++;
        }
    }

    //cazul COREP - 17.2 - OPR DETAILS 2 - dinamic (nr variabil de randuri)
    public static void writeXmlCorep17_2()
    {
        int currentDynamicRow = 1;

        while(getCellValue(sheet, tableGroupfieldColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String groupfieldValue = getCellValue(sheet, tableGroupfieldColumn + (startRow + currentDynamicRow));

            String rowCode, xmlRowCode = "";

            xmlRowCode = xmlRowNamePattern;

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {
                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                    currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                    //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                    currentXmlTag = currentXmlTag.replace("@I056", valueOfI056);

                    currentXmlTag = currentXmlTag.replace("@GROUPFIELD", groupfieldValue);


                    if (colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0040"))
                    {
                        currentXmlTag = currentXmlTag.replace("allocated_corep_text", "allocated_corep_date");
                        currentXmlTag = currentXmlTag.replace("TEXT=\"@TEXT\"", "DATEVAL=\"@DATEVAL\"");
                    }

                    currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                    currentXmlTag = currentXmlTag.replace("@DATEVAL", cellValue);

                    writeXml(currentXmlTag);
                }
            }

            currentDynamicRow++;
        }
    }

    //cazul COREP - 32.3 - PRUVAL 3 - dinamic (nr variabil de randuri)
    public static void writeXmlCorep32_3()
    {
        int currentDynamicRow = 1;

        while(getCellValue(sheet, tableRankColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String rankValue = getFormattedCellValue(getCellValue(sheet, tableRankColumn + (startRow + currentDynamicRow)));

            String rowCode, xmlRowCode = "";

            xmlRowCode = xmlRowNamePattern;

            String currentXmlTag = xmlPattern;

            if (currentXmlTag != null)
            {

                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                currentXmlTag = currentXmlTag.replace("@colname", "DUMMY_DYN_CELL");
                currentXmlTag = currentXmlTag.replace("@rowname", "d.x");
                currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");
                currentXmlTag = currentXmlTag.replace("@RANK", rankValue);

                writeXml(currentXmlTag);
            }

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {
                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {
                    if(!colCode.equals("0005"))
                    {
                        if(colCode.equals("0010") || colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0040"))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                            //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                            if(cellValue.startsWith("("))
                                cellValue = getFormattedCellValue(cellValue);
                            currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@RANK", rankValue);
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                            //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                            currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");
                            currentXmlTag = currentXmlTag.replace("@RANK", rankValue);
                        }

                        writeXml(currentXmlTag);
                    }
                }
            }

            currentDynamicRow++;
        }
    }

    //cazul COREP - 32.4 - PRUVAL 4 - dinamic (nr variabil de randuri)
    public static void writeXmlCorep32_4()
    {
        int currentDynamicRow = 1;

        while(getCellValue(sheet, tableRankColumn + (startRow + currentDynamicRow)) != "")
        {
            Row row = sheet.getRow(startRow + currentDynamicRow - 1);

            String rankValue = getFormattedCellValue(getCellValue(sheet, tableRankColumn + (startRow + currentDynamicRow)));

            String rowCode, xmlRowCode = "";

            xmlRowCode = xmlRowNamePattern;

            String currentXmlTag = xmlPattern;

            if (currentXmlTag != null)
            {

                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                currentXmlTag = currentXmlTag.replace("@colname", "DUMMY_DYN_CELL");
                currentXmlTag = currentXmlTag.replace("@rowname", "d.x");
                currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");
                currentXmlTag = currentXmlTag.replace("@RANK", rankValue);

                writeXml(currentXmlTag);
            }

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {
                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {
                    if(!colCode.equals("0005"))
                    {
                        if(colCode.equals("0010") || colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0050"))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                            //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                            if(cellValue.startsWith("("))
                                cellValue = getFormattedCellValue(cellValue);
                            currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@RANK", rankValue);
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");
                            //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                            //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale

                            currentXmlTag = currentXmlTag.replace("@RANK", rankValue);
                        }

                        writeXml(currentXmlTag);
                    }
                }
            }

            currentDynamicRow++;
        }
    }

    //cazul sheet-urilor multiplicate de user, prefixul sheet-ului ramane cel din baza de date
    //----------------------------------------------------------------------------------------
    //cazul COREP - 9.1 - CR GB 1 - Geographical Breakdown SA Exposures - sheet multiplicat dinamic
    //cazul COREP - 9.2 - CR GB 2 - Geographical Breakdown IRB Exposures - sheet multiplicat dinamic
    //cazul COREP - 9.4 - CR GB 4 - Geographical Breakdown countercyclical (CCB) - sheet multiplicat dinamic
    //cazul COREP - 18 - MKR TDI - Traded Debt Instruments - sheet multiplicat dinamic
    //cazul COREP - 21 - MKR EQU - Equities - sheet multiplicat dinamic
    //cazul COREP - 33 - GENERAL GOV - Equities - sheet multiplicat dinamic
    //cazul IP_LOSSES - 15 - sheet multiplicat dinamic
    public static void writeXmlCorepDynamicSheet()
    {
        //caut sheet-urile dinamice
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            sheet = workbook.getSheetAt(sheetIndex);

            if (sheet.getSheetName().startsWith(currentSheetName))
            {
                valueOfB272 = "";

                if(tableB272Cell != null)
                    valueOfB272 = getCellValue(sheet, tableB272Cell);

                valueOfB271 = "";

                if(tableB271Cell != null)
                    valueOfB271 = getCellValue(sheet, tableB271Cell);

                for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);

                    String rowCode, xmlRowCode = "";

                    if (tableRowCodeColumn.equals("*"))
                    {
                        rowCode = "010";
                        xmlRowCode = "010";
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1));
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);

                        if (currentSheetName.equals("9.4") && valueOfB272.equals("TOTAL"))
                            xmlRowCode = xmlRowCode.replace("d4.", "");
                        else if (currentSheetName.equals("18") && (valueOfB271.equals("TOT") || valueOfB271.equals("OTH")))
                            xmlRowCode = xmlRowCode.replace("d.", "");
                        else if (currentSheetName.equals("21") && (valueOfB272.equals("WP") || valueOfB272.equals("OT")))
                            xmlRowCode = xmlRowCode.replace("d.", "");
                        else if (currentSheetName.equals("33") && valueOfB272.equals("WP"))
                            xmlRowCode = xmlRowCode.replace("d2.XD", "X");
                    }

                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        if(isIgnoredColor(cell))
                            continue;

                        String colCode, xmlColCode = "";

                        if (tableColumnCodeRow.equals("*"))
                        {
                            colCode = "010";
                            xmlColCode = "010";
                        }
                        else
                        {
                            colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                            xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                        }

                        String cellValue = getCellValue(cell);

                        if (cellValue.equals(""))
                            continue;

                        if(currentSheetName.equals("9.4") && cellValue.equals("0"))
                            continue;

                        String currentXmlTag = xmlPattern;

                        if (currentXmlTag != null)
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);


                            if (currentSheetName.equals("18") && (valueOfB271.equals("TOT") || valueOfB271.equals("OTH")))
                            {
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName.replace("alloc", "sum_alloc"));
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName.replace("alloc", "sum_alloc"));
                            }
                            else if (currentSheetName.equals("21") && (valueOfB272.equals("WP") || valueOfB272.equals("OT")))
                            {
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName.replace("alloc", "sum_alloc"));
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName.replace("alloc", "sum_alloc"));
                            }
                            else if (currentSheetName.equals("33") && valueOfB272.equals("WP"))
                            {
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName.replace("alloc", "sum_alloc"));
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName.replace("alloc", "sum_alloc"));
                            }
                            else
                            {
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            }

                            //if(currentSheetName.startsWith("15"))
                            //	System.out.println(reportName);


                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                            //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale



                            if (currentSheetName.equals("18"))
                            {
                                if (valueOfB271.equals("TOT"))
                                    currentXmlTag = currentXmlTag.replace("@I056", "996");
                                else if (valueOfB271.equals("OTH"))
                                    currentXmlTag = currentXmlTag.replace("@I056", "995");
                                else
                                    currentXmlTag = currentXmlTag.replace(" I056=\"@I056\"", "");
                            }
                            else if (currentSheetName.equals("21"))
                            {
                                if (valueOfB272.equals("WP"))
                                    currentXmlTag = currentXmlTag.replace("@I056", "994");
                                else if (valueOfB272.equals("OT"))
                                    currentXmlTag = currentXmlTag.replace("@I056", "993");
                                else
                                    currentXmlTag = currentXmlTag.replace(" I056=\"@I056\"", "");
                            }


                            if (valueOfB272.equals("TOTAL") &&
                                    (currentSheetName.equals("9.1") ||
                                            currentSheetName.equals("9.2") ||
                                            currentSheetName.equals("9.4")))
                                currentXmlTag = currentXmlTag.replace(" B272=\"@B272\"", "");
                            else if (valueOfB272.equals("WP") && currentSheetName.equals("33"))
                                currentXmlTag = currentXmlTag.replace(" B272=\"@B272\"", "");
                            else
                                currentXmlTag = currentXmlTag.replace("@B272", valueOfB272);


                            currentXmlTag = currentXmlTag.replace("@B271", valueOfB271);


                            currentXmlTag = currentXmlTag.replace("@B790", "#B790");
                            currentXmlTag = currentXmlTag.replace("@C309", "#C309");
                            currentXmlTag = currentXmlTag.replace("@RANK", "#RANK");
                            currentXmlTag = currentXmlTag.replace("@GROUPFIELD", "#GROUPFIELD");


                            if(currentXmlTag.indexOf("allocated_regulator_text") != -1)
                                currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                            else
                                currentXmlTag = currentXmlTag.replace("@TEXT", "#TEXT");


                            if ((xlsFileId.equals("1") || xlsFileName.equals("Annex_1_(Solvency).xlsx")) && currentSheetName.startsWith("8"))
                            {
                                if (tableB015Value != null)
                                    currentXmlTag = currentXmlTag.replace("@B015", tableB015Value);
                                else
                                    currentXmlTag = currentXmlTag.replace(" B015=\"@B015\"", "");

                                if(currentSheetName.startsWith("8.1") && xmlRowCode.equals("015"))
                                    currentXmlTag = currentXmlTag.replace("@B046", "1");
                                else
                                    currentXmlTag = currentXmlTag.replace(" B046=\"@B046\"", "");
                            }


                            writeXml(currentXmlTag);
                        }
                    }
                }
            }
        }
    }


    //sheet-uri default corep
    public static void writeXmlCorepDefault()
    {
        if(sheet != null)
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1));
                    xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.equals(""))
                        continue;

                    if(currentSheetName.startsWith("7_") && cellValue.equals("0"))
                        continue;

                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        //currentXmlTag = currentXmlTag.replace("@P024", "0"); //optionale
                        //currentXmlTag = currentXmlTag.replace("@J073", "ALL"); //optionale
                        currentXmlTag = currentXmlTag.replace("@I056", valueOfI056);

                        currentXmlTag = currentXmlTag.replace("@B272", valueOfB272);
                        currentXmlTag = currentXmlTag.replace("@B271", "#B271");
                        currentXmlTag = currentXmlTag.replace("@B790", "#B790");


                        if ((xlsFileId.equals("1") || xlsFileName.equals("Annex_1_(Solvency).xlsx")) && currentSheetName.equals("22"))
                            if(rowCode.equals("130"))
                                currentXmlTag = currentXmlTag.replace("@C309", "EUR");
                            else if(rowCode.equals("140"))
                                currentXmlTag = currentXmlTag.replace("@C309", "ALL");
                            else if(rowCode.equals("150"))
                                currentXmlTag = currentXmlTag.replace("@C309", "ARS");
                            else if(rowCode.equals("160"))
                                currentXmlTag = currentXmlTag.replace("@C309", "AUD");
                            else if(rowCode.equals("170"))
                                currentXmlTag = currentXmlTag.replace("@C309", "BRL");
                            else if(rowCode.equals("180"))
                                currentXmlTag = currentXmlTag.replace("@C309", "BGN");
                            else if(rowCode.equals("190"))
                                currentXmlTag = currentXmlTag.replace("@C309", "CAD");
                            else if(rowCode.equals("200"))
                                currentXmlTag = currentXmlTag.replace("@C309", "CZK");
                            else if(rowCode.equals("210"))
                                currentXmlTag = currentXmlTag.replace("@C309", "DKK");
                            else if(rowCode.equals("220"))
                                currentXmlTag = currentXmlTag.replace("@C309", "EGP");
                            else if(rowCode.equals("230"))
                                currentXmlTag = currentXmlTag.replace("@C309", "GBP");
                            else if(rowCode.equals("240"))
                                currentXmlTag = currentXmlTag.replace("@C309", "HUF");
                            else if(rowCode.equals("250"))
                                currentXmlTag = currentXmlTag.replace("@C309", "JPY");
                            else if(rowCode.equals("270"))
                                currentXmlTag = currentXmlTag.replace("@C309", "LVL");
                            else if(rowCode.equals("280"))
                                currentXmlTag = currentXmlTag.replace("@C309", "MKD");
                            else if(rowCode.equals("290"))
                                currentXmlTag = currentXmlTag.replace("@C309", "MXN");
                            else if(rowCode.equals("300"))
                                currentXmlTag = currentXmlTag.replace("@C309", "PLN");
                            else if(rowCode.equals("310"))
                                currentXmlTag = currentXmlTag.replace("@C309", "RON");
                            else if(rowCode.equals("320"))
                                currentXmlTag = currentXmlTag.replace("@C309", "RUB");
                            else if(rowCode.equals("330"))
                                currentXmlTag = currentXmlTag.replace("@C309", "RSD");
                            else if(rowCode.equals("340"))
                                currentXmlTag = currentXmlTag.replace("@C309", "SEK");
                            else if(rowCode.equals("350"))
                                currentXmlTag = currentXmlTag.replace("@C309", "CHF");
                            else if(rowCode.equals("360"))
                                currentXmlTag = currentXmlTag.replace("@C309", "TRY");
                            else if(rowCode.equals("370"))
                                currentXmlTag = currentXmlTag.replace("@C309", "UAH");
                            else if(rowCode.equals("380"))
                                currentXmlTag = currentXmlTag.replace("@C309", "USD");
                            else if(rowCode.equals("390"))
                                currentXmlTag = currentXmlTag.replace("@C309", "ISK");
                            else if(rowCode.equals("400"))
                                currentXmlTag = currentXmlTag.replace("@C309", "NOK");
                            else if(rowCode.equals("410"))
                                currentXmlTag = currentXmlTag.replace("@C309", "HKD");
                            else if(rowCode.equals("420"))
                                currentXmlTag = currentXmlTag.replace("@C309", "TWD");
                            else if(rowCode.equals("430"))
                                currentXmlTag = currentXmlTag.replace("@C309", "NZD");
                            else if(rowCode.equals("440"))
                                currentXmlTag = currentXmlTag.replace("@C309", "SGD");
                            else if(rowCode.equals("450"))
                                currentXmlTag = currentXmlTag.replace("@C309", "KRW");
                            else if(rowCode.equals("460"))
                                currentXmlTag = currentXmlTag.replace("@C309", "CNY");
                            else if(rowCode.equals("470"))
                                currentXmlTag = currentXmlTag.replace("@C309", "OTH");
                            else if(rowCode.equals("480"))
                                currentXmlTag = currentXmlTag.replace("@C309", "HRK");
                            else
                                currentXmlTag = currentXmlTag.replace(" C309=\"@C309\"", "");
                        else
                            currentXmlTag = currentXmlTag.replace("@C309", "#C309");


                        currentXmlTag = currentXmlTag.replace("@RANK", "#RANK");
                        currentXmlTag = currentXmlTag.replace("@GROUPFIELD", "#GROUPFIELD");


                        if(currentXmlTag.indexOf("allocated_regulator_text") != -1)
                            currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                        else
                            currentXmlTag = currentXmlTag.replace("@TEXT", "#TEXT");


                        if((xlsFileId.equals("1") || xlsFileName.equals("Annex_1_(Solvency).xlsx")) && currentSheetName.startsWith("8"))
                        {
                            if(valueOfI056.trim().equals("998") || valueOfI056.trim().equals("10") || valueOfI056.trim().equals("20") ||
                                    valueOfI056.trim().equals("31") || valueOfI056.trim().equals("997") || valueOfI056.trim().equals("42") ||
                                    valueOfI056.trim().equals("110") || valueOfI056.trim().equals("120") || valueOfI056.trim().equals("131"))
                            {
                                currentXmlTag = currentXmlTag.replace(" B015=\"@B015\"", "");
                                currentXmlTag = currentXmlTag.replace(" B046=\"@B046\"", "");
                            }
                            else if(valueOfI056.trim().equals("30") || valueOfI056.trim().equals("41") || valueOfI056.trim().equals("40") ||
                                    valueOfI056.trim().equals("130"))
                            {
                                currentXmlTag = currentXmlTag.replace("@B015", tableB015Value);
                                if(currentSheetName.startsWith("8.1") && xmlRowCode.equals("015") && tableB015Value.equals("1"))
                                    currentXmlTag = currentXmlTag.replace("@B046", "1");
                                else
                                    currentXmlTag = currentXmlTag.replace(" B046=\"@B046\"", "");
                            }
							/*else
							{
								if(tableB015Value != null)
									currentXmlTag = currentXmlTag.replace("@B015", tableB015Value);
								else
									currentXmlTag = currentXmlTag.replace(" B015=\"@B015\"", "");

								if(currentSheetName.startsWith("8.1") && xmlRowCode.equals("015"))
									currentXmlTag = currentXmlTag.replace("@B046", "1");
								else
									currentXmlTag = currentXmlTag.replace(" B046=\"@B046\"", "");
							}*/
                        }


                        writeXml(currentXmlTag);
                    }
                }
            }
        }
    }


    //sheet-uri default NSFR
    public static void writeXmlNsfrDefault()
    {
        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);

            String rowCode, xmlRowCode = "";

            if (tableRowCodeColumn.equals("*"))
            {
                rowCode = "010";
                xmlRowCode = "010";
            }
            else
            {
                if(xmlRowNamePattern.equals("ABACUS"))
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                    if(abacusRowNames.containsKey(rowCode))
                        xmlRowCode = abacusRowNames.get(rowCode);
                }
                else
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                    xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                }
            }

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {
                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                    currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@C007", "RON");
                    currentXmlTag = currentXmlTag.replace("@J073", "RON");


                    writeXml(currentXmlTag);
                }
            }
        }
    }


    //sheet-uri default NSFR Valuta
    public static void writeXmlNsfrCcyDefault()
    {

        //caut sheet-urile dinamice
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            sheet = workbook.getSheetAt(sheetIndex);

            valueOfC007 = "";

            if (sheet.getSheetName().startsWith(currentSheetName))
            {

                valueOfC007 = getCellValue(sheet, tableC007Cell);

                for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);

                    String rowCode, xmlRowCode = "";

                    if (tableRowCodeColumn.equals("*"))
                    {
                        rowCode = "010";
                        xmlRowCode = "010";
                    }
                    else
                    {
                        if(xmlRowNamePattern.equals("ABACUS"))
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            if(abacusRowNames.containsKey(rowCode))
                                xmlRowCode = abacusRowNames.get(rowCode);
                        }
                        else
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                        }
                    }

                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        if(isIgnoredColor(cell))
                            continue;

                        String colCode, xmlColCode = "";

                        if (tableColumnCodeRow.equals("*"))
                        {
                            colCode = "010";
                            xmlColCode = "010";
                        }
                        else
                        {

                            colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                            xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                        }

                        String cellValue = getCellValue(cell);

                        if (cellValue.equals(""))
                            continue;

                        String currentXmlTag = xmlPattern;

                        if (currentXmlTag != null)
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");

                            if(valueOfC007.equals("RON_EQ"))
                                currentXmlTag =  currentXmlTag.replace("C007=\"@C007\" ", "");
                            else
                                currentXmlTag = currentXmlTag.replace("@C007", valueOfC007);
                            if(valueOfC007.equals("RON_EQ"))
                                currentXmlTag = currentXmlTag.replace("@J073", "Local Currency");
                            else
                                currentXmlTag = currentXmlTag.replace("@J073", valueOfC007);

                            writeXml(currentXmlTag);
                        }
                    }
                }
            }
        }
    }


    //sheet-uri default LCR
    public static void writeXmlLcrDefault()
    {

        if(currentSheetName.equals("77"))
        {
            int counter = 0;
            int lcrLiqCategoryVal = 1;
            String valF003 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@colname", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@rowname", "d.line");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", "0.00");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@F003",lcrLiqCategoryVal + "");
                    currentXmlTag = currentXmlTag.replace("@LCR_LIQ_CATEGORY", lcrLiqCategoryVal + "");
                    currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    //if(colCode.equals("005"))
                    //	valF003 = getCellValue(row.getCell(columnIndex + 2));//col 0020


                    if (currentXmlTag != null) // && !colCode.equals("020")
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@rowname", xmlRowNamePattern);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", "0");
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@F003",lcrLiqCategoryVal + "");
                        currentXmlTag = currentXmlTag.replace("@LCR_LIQ_CATEGORY", lcrLiqCategoryVal + "");
                        currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);



                        writeXml(currentXmlTag);
                    }
                }
                lcrLiqCategoryVal++;
                row = sheet.getRow(startRow + counter);
            }

        }
        else
        {


            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell ;

                    try
                    {
                        cell = row.getCell(columnIndex);
                    }
                    catch(Exception ex)
                    {
                        continue;
                    }


                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.equals(""))
                        continue;

                    if((currentSheetName.startsWith("72") || currentSheetName.startsWith("73") || currentSheetName.startsWith("76") || currentSheetName.startsWith("74")) && cellValue.equals("0"))
                        continue;

                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@C007", "RON");
                        currentXmlTag = currentXmlTag.replace("@J073", "RON");

                        writeXml(currentXmlTag);
                    }
                }
            }
        }
    }


    //sheet-uri default LCR Valuta
    public static void writeXmlLcrCcyDefault()
    {
        if(currentSheetName.equals("77"))
        {
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);
                int counter = 0;
                int lcrLiqCategoryVal = 1;
                String valF003 = "";

                if (sheet.getSheetName().startsWith(currentSheetName))
                {
                    Row row = sheet.getRow(startRow);



                    while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
                    {
                        counter++;

                        String currentXmlTag = xmlPattern;

                        if (currentXmlTag != null)
                        {

                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", "DUMMY_DYN_CELL");
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.line");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", "0.00");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@F003",lcrLiqCategoryVal + "");//reportingEntity
                            currentXmlTag = currentXmlTag.replace("@LCR_LIQ_CATEGORY", lcrLiqCategoryVal + "");
                            currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");

                            writeXml(currentXmlTag);
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {
                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.trim().equals(""))
                                continue;

                            currentXmlTag = xmlPattern;

                            //if(colCode.equals("005"))
                            //	valF003 = getCellValue(row.getCell(columnIndex + 2));//col 0020


                            if (currentXmlTag != null) // && !colCode.equals("020")
                            {
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@rowname", xmlRowNamePattern);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", "0");
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                currentXmlTag = currentXmlTag.replace("@F003",lcrLiqCategoryVal + ""); //reportingEntity
                                currentXmlTag = currentXmlTag.replace("@LCR_LIQ_CATEGORY", lcrLiqCategoryVal + "");
                                currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);



                                writeXml(currentXmlTag);
                            }
                        }

                        lcrLiqCategoryVal++;
                        row = sheet.getRow(startRow + counter);
                    }
                }
            }
        }
        else
        {
            //caut sheet-urile dinamice
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfC007 = "";

                if (sheet.getSheetName().startsWith(currentSheetName))
                {

                    valueOfC007 = getCellValue(sheet, tableC007Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell;

                            try
                            {
                                cell = row.getCell(columnIndex);
                            }
                            catch(Exception ex)
                            {
                                continue;
                            }

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.equals(""))
                                continue;

                            if((currentSheetName.startsWith("72") || currentSheetName.startsWith("73") || currentSheetName.startsWith("76") || currentSheetName.startsWith("74")) && cellValue.equals("0"))
                                continue;

                            String currentXmlTag = xmlPattern;

                            if (currentXmlTag != null)
                            {
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                if(valueOfC007.equals("RON_EQ"))
                                    currentXmlTag =  currentXmlTag.replace("C007=\"@C007\" ", "");
                                else
                                    currentXmlTag = currentXmlTag.replace("@C007", valueOfC007);
                                if(valueOfC007.equals("RON_EQ"))
                                    currentXmlTag = currentXmlTag.replace("@J073", "Local Currency");
                                else
                                    currentXmlTag = currentXmlTag.replace("@J073", valueOfC007);

                                writeXml(currentXmlTag);
                            }
                        }
                    }
                }
            }
        }
    }


    //sheet-uri default ALMM
    public static void writeXmlAlmmDefault()
    {
        String [] c67rows = new String[]{"020","030","040","050","060","070","080","090","100","110"};
        String [] c67columns = new String[] {"010","020","030","040","050"};
        String [] c71rows = new String[]{"020","030","040","050","060","070","080","090","100","110"};
        String [] c71columns = new String[] {"010","020","030","040","050","060","070"};


        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);

            String rowCode, xmlRowCode = "";

            if (tableRowCodeColumn.equals("*"))
            {
                rowCode = "010";
                xmlRowCode = "010";
            }
            else
            {
                if(xmlRowNamePattern.equals("ABACUS"))
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                    if(abacusRowNames.containsKey(rowCode))
                        xmlRowCode = abacusRowNames.get(rowCode);
                }
                else
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                    xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                }
            }

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;



                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    if(sheet.getSheetName().startsWith("67") && Arrays.asList(c67rows).contains(rowCode) && Arrays.asList(c67columns).contains(colCode))
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@F001CALC", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("colname=\"@colname\"", "COLNAME=\"" + xmlColCode + "\"");
                        currentXmlTag = currentXmlTag.replace("rowname=\"@rowname\"", "ROWNAME=\"" + xmlRowCode + "\"");
                        currentXmlTag = currentXmlTag.replace("@TEXT", getFormattedCellValue(cellValue));
                        currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                        currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                        currentXmlTag = currentXmlTag.replace("@C007", "RON");
                        currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                        currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                        currentXmlTag = currentXmlTag.replace(" AMOUNT=\"@AMOUNT\"", "");
                        currentXmlTag = currentXmlTag.replace("allocated_almm","ALLOCATED_ALMM_TEXT");

                    }
                    else if (sheet.getSheetName().startsWith("71") && Arrays.asList(c71rows).contains(rowCode) && Arrays.asList(c71columns).contains(colCode))
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@F001CALC", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("colname=\"@colname\"", "COLNAME=\"" + xmlColCode + "\"");
                        currentXmlTag = currentXmlTag.replace("rowname=\"@rowname\"", "ROWNAME=\"" + xmlRowCode + "\"");
                        currentXmlTag = currentXmlTag.replace("@TEXT", getFormattedCellValue(cellValue));
                        currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                        currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                        currentXmlTag = currentXmlTag.replace("@C007", "RON");
                        currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                        currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                        currentXmlTag = currentXmlTag.replace(" AMOUNT=\"@AMOUNT\"", "");
                        currentXmlTag = currentXmlTag.replace("allocated_almm","ALLOCATED_ALMM_TEXT");
                    }
                    else
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@C007", "RON");
                        currentXmlTag = currentXmlTag.replace("@J073", "RON");
                        currentXmlTag = currentXmlTag.replace(" F001CALC=\"@F001CALC\"", "");
                        currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");

                    }

                    writeXml(currentXmlTag);
                }
            }
        }

    }


    //sheet-uri default ALMM Valuta
    public static void writeXmlAlmmCcyDefault()
    {
        String [] c67rows = new String[]{"020","030","040","050","060","070","080","090","100","110"};
        String [] c67columns = new String[] {"010","020","030","040","050"};
        String [] c71rows = new String[]{"020","030","040","050","060","070","080","090","100","110"};
        String [] c71columns = new String[] {"010","020","030","040","050","060","070"};

        //Arrays.asList(yourArray).contains(yourValue);

        //caut sheet-urile dinamice
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            sheet = workbook.getSheetAt(sheetIndex);

            valueOfC007 = "";

            if (sheet.getSheetName().startsWith(currentSheetName))
            {

                valueOfC007 = getCellValue(sheet, tableC007Cell);

                for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);

                    String rowCode, xmlRowCode = "";

                    if (tableRowCodeColumn.equals("*"))
                    {
                        rowCode = "010";
                        xmlRowCode = "010";
                    }
                    else
                    {
                        if(xmlRowNamePattern.equals("ABACUS"))
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            if(abacusRowNames.containsKey(rowCode))
                                xmlRowCode = abacusRowNames.get(rowCode);
                        }
                        else
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                        }
                    }

                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        if(isIgnoredColor(cell))
                            continue;

                        String colCode, xmlColCode = "";

                        if (tableColumnCodeRow.equals("*"))
                        {
                            colCode = "010";
                            xmlColCode = "010";
                        }
                        else
                        {

                            colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                            xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                        }

                        String cellValue = getCellValue(cell);

                        if (cellValue.equals(""))
                            continue;



                        String currentXmlTag = xmlPattern;

                        if (currentXmlTag != null)
                        {

                            if(sheet.getSheetName().startsWith("67") && Arrays.asList(c67rows).contains(rowCode) && Arrays.asList(c67columns).contains(colCode))
                            {
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@F001CALC", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("colname=\"@colname\"", "COLNAME=\"" + xmlColCode + "\"");
                                currentXmlTag = currentXmlTag.replace("rowname=\"@rowname\"", "ROWNAME=\"" + xmlRowCode + "\"");
                                currentXmlTag = currentXmlTag.replace("@TEXT", getFormattedCellValue(cellValue));
                                currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                                currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                                if(valueOfC007.equals("RON_EQ"))
                                    currentXmlTag =  currentXmlTag.replace("C007=\"@C007\" ", "");
                                else
                                    currentXmlTag = currentXmlTag.replace("@C007", valueOfC007);
                                currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                                currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                                currentXmlTag = currentXmlTag.replace(" AMOUNT=\"@AMOUNT\"", "");
                                currentXmlTag = currentXmlTag.replace("allocated_almm","ALLOCATED_ALMM_TEXT");

                            }
                            else if (sheet.getSheetName().startsWith("71") && Arrays.asList(c71rows).contains(rowCode) && Arrays.asList(c71columns).contains(colCode))
                            {
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@F001CALC", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("colname=\"@colname\"", "COLNAME=\"" + xmlColCode + "\"");
                                currentXmlTag = currentXmlTag.replace("rowname=\"@rowname\"", "ROWNAME=\"" + xmlRowCode + "\"");
                                currentXmlTag = currentXmlTag.replace("@TEXT", getFormattedCellValue(cellValue));
                                currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                                currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                                if(valueOfC007.equals("RON_EQ"))
                                    currentXmlTag =  currentXmlTag.replace("C007=\"@C007\" ", "");
                                else
                                    currentXmlTag = currentXmlTag.replace("@C007", valueOfC007);
                                currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                                currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                                currentXmlTag = currentXmlTag.replace(" AMOUNT=\"@AMOUNT\"", "");
                                currentXmlTag = currentXmlTag.replace("allocated_almm","ALLOCATED_ALMM_TEXT");
                            }
                            else
                            {
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                if(valueOfC007.equals("RON_EQ"))
                                    currentXmlTag =  currentXmlTag.replace("C007=\"@C007\" ", "");
                                else
                                    currentXmlTag = currentXmlTag.replace("@C007", valueOfC007);
                                if(valueOfC007.equals("RON_EQ"))
                                    currentXmlTag = currentXmlTag.replace("@J073", "Local Currency");
                                else
                                    currentXmlTag = currentXmlTag.replace("@J073", valueOfC007);
                                currentXmlTag = currentXmlTag.replace(" F001CALC=\"@F001CALC\"", "");
                                currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");

                            }

                            writeXml(currentXmlTag);
                        }
                    }
                }
            }
        }
    }


    //sheet-uri default Large Exposures
    public static void writeXmlLeDefault()
    {

        if(currentSheetName.equals("26"))
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.equals(""))
                        continue;



                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {

                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@C007", "RON");
                        currentXmlTag = currentXmlTag.replace("@J073", "EUR");

                        writeXml(currentXmlTag);
                    }
                }
            }
        }
        //sheets dinamice
        else
        {
            int counter = 0;
            String incValue = "";
            String gccValue = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                //Row row = sheet.getRow(rowIndex);
                counter++;

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }


                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {
                    if(!currentSheetName.equals("26"))
                        incValue = getCellValue(row.getCell(stringToInt.get(startColumn)));
                    if(currentSheetName.equals("29") || currentSheetName.equals("31"))
                        gccValue = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@colname", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@rowname", "d.A");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0100000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@C007", "RON");
                    currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                    currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");
                    if(!incValue.equals(""))
                        currentXmlTag = currentXmlTag.replace("@INC", incValue);
                    else
                        currentXmlTag = currentXmlTag.replace(" INC=\"@INC\"", "");
                    if(!gccValue.equals(""))
                        currentXmlTag = currentXmlTag.replace("@GCC", gccValue);
                    else
                        currentXmlTag = currentXmlTag.replace(" GCC=\"@GCC\"", "");

                    incValue = "";
                    gccValue = "";

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if((currentSheetName.equals("28") && colCode.equals("010")) || (currentSheetName.equals("27") && colCode.equals("010")) )
                        incValue = cellValue;
                    else if(currentSheetName.equals("29") && colCode.equals("010"))
                        incValue = cellValue;
                    else if(currentSheetName.equals("29") && colCode.equals("020"))
                        gccValue = cellValue;
                    else if((currentSheetName.equals("30") && colCode.equals("010")) || currentSheetName.equals("31") && colCode.equals("010"))
                        incValue = cellValue;
                    else if(currentSheetName.equals("31") && colCode.equals("020"))
                        gccValue = cellValue;


                    if (currentXmlTag != null)
                    {
                        if((currentSheetName.equals("28") && !colCode.equals("010") && !colCode.equals("020") && !colCode.equals("030")) ||
                                (currentSheetName.equals("29") && !colCode.equals("010") && !colCode.equals("020") && !colCode.equals("030") && !colCode.equals("040")) ||
                                (currentSheetName.equals("30") && !colCode.equals("010")) ||
                                (currentSheetName.equals("31") && !colCode.equals("010") && !colCode.equals("020")))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.A");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@C007", "RON");
                            currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                            currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", "");
                            if(!incValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INC", incValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INC=\"@INC\"", "");
                            if(!gccValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@GCC", gccValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" GCC=\"@GCC\"", "");
                        }

                        else if((currentSheetName.equals("28") || currentSheetName.equals("29")) && colCode.equals("030"))
                        {
                            currentXmlTag = currentXmlTag.replace("allocated_le", "allocated_le_int");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.A");
                            currentXmlTag = currentXmlTag.replace(" TEXT=\"@TEXT\"", " VALUE=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace(" AMOUNT=\"@AMOUNT\"", "");
                            currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                            currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                            currentXmlTag = currentXmlTag.replace(" C007=\"@C007\"", "");
                            currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                            currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                            if(!incValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INC", incValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INC=\"@INC\"", "");
                            if(!gccValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@GCC", gccValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" GCC=\"@GCC\"", "");
                        }
                        else
                        {
                            if((currentSheetName.equals("27") && colCode.equals("010")) ||
                                    (currentSheetName.equals("28") && colCode.equals("010")) ||
                                    (currentSheetName.equals("29") && (colCode.equals("010") || colCode.equals("020"))) ||
                                    (currentSheetName.equals("30") && colCode.equals("010")) ||
                                    (currentSheetName.equals("31") && (colCode.equals("010") || colCode.equals("020"))))
                                continue;
                            currentXmlTag = currentXmlTag.replace("allocated_le", "allocated_le_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.A");
                            currentXmlTag = currentXmlTag.replace("@TEXT", cellValue);
                            currentXmlTag = currentXmlTag.replace(" AMOUNT=\"@AMOUNT\"", "");
                            currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                            currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                            currentXmlTag = currentXmlTag.replace(" C007=\"@C007\"", "");
                            currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                            currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                            if(!incValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INC", incValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INC=\"@INC\"", "");
                            if(!gccValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@GCC", gccValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" GCC=\"@GCC\"", "");
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }

    }


    //sheet-uri default AE
    public static void writeXmlAeDefault()
    {
        String valueOfB271 = "";

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            sheet = workbook.getSheetAt(sheetIndex);

            valueOfC200 = "";

            if (sheet.getSheetName().startsWith(currentSheetName))
            {
                if(currentSheetName.startsWith("35"))
                    valueOfC200 = getCellValue(sheet, tableC200Cell);

                for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);

                    String rowCode, xmlRowCode = "";

                    if (tableRowCodeColumn.equals("*"))
                    {
                        rowCode = "010";
                        xmlRowCode = "010";
                    }
                    else
                    {
                        if(xmlRowNamePattern.equals("ABACUS"))
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            if(abacusRowNames.containsKey(rowCode))
                                xmlRowCode = abacusRowNames.get(rowCode);
                        }
                        else
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                        }
                    }

                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        if(isIgnoredColor(cell))
                            continue;

                        String colCode, xmlColCode = "";

                        if (tableColumnCodeRow.equals("*"))
                        {
                            colCode = "010";
                            xmlColCode = "010";
                        }
                        else
                        {

                            colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                            xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                        }

                        String cellValue = getCellValue(cell);

                        if (cellValue.equals(""))
                            continue;

                        String currentXmlTag = xmlPattern;

                        if (currentXmlTag != null)
                        {
                            if(currentSheetName.equals("35") && (colCode.equals("010") || colCode.equals("012") || colCode.equals("090") || colCode.equals("100") ||
                                    colCode.equals("110") || colCode.equals("120") || colCode.equals("130") || colCode.equals("140")))
                            {
                                currentXmlTag = currentXmlTag.replace("reportdata", "allocated_ae_text");
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                                currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                //currentXmlTag = currentXmlTag.replace("@C007", "RON");
                                currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                                currentXmlTag = currentXmlTag.replace("@C200", valueOfC200);
                                currentXmlTag = currentXmlTag.replace("@AE910", valueOfC200);
                            }
                            else
                            {
                                if(currentSheetName.equals("34") && (colCode.equals("010") || colCode.equals("020")))
                                {
                                    xmlColCode = xmlColCode.replace(" / d.x","");
                                    currentXmlTag = currentXmlTag.replace(" B271=\"@B271\"", "");
                                }
                                else if(currentSheetName.equals("34"))
                                {
                                    valueOfB271 = "";

                                    valueOfB271 = getCellValue(sheet, intToString.get(columnIndex) + (Integer.parseInt(tableColumnCodeRow) - 1));
                                    currentXmlTag = currentXmlTag.replace("@B271", valueOfB271);
                                    xmlColCode = "d.x";

                                    if(valueOfB271.trim().equals(""))
                                        continue;
                                }

                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                                currentXmlTag = currentXmlTag.replace("@C200", valueOfC200);
                                currentXmlTag = currentXmlTag.replace("@AE910", valueOfC200);
                            }


                            writeXml(currentXmlTag);
                        }
                    }
                }
            }
        }
    }

    //sheet-uri default Rezolutie
    public static void writeXmlResolutionDefault()
    {

        if(currentSheetName.equals("1 ORG"))
        {
            int counter = 0;
            String reportDataCols [] = new String[] {"0090", "0100", "0110", "0130", "0140", "0150", "0200", "0210"};
            String entityCode = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                entityCode = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@ENTITY_CODE", entityCode);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                        entityCode = getCellValue(row.getCell(columnIndex + 1));

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020"))
                            continue;
                        if(Arrays.asList(reportDataCols).contains(colCode))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@ENTITY_CODE", entityCode);

                        }
                        else if(colCode.equals("0040"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_int");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "VALUE=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@ENTITY_CODE", entityCode);
                        }
                        else
                        {
                            if(colCode.equals("0050"))
                                cellValue = getFormattedCellValue(cellValue);
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@ENTITY_CODE", entityCode);
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }

        }
        else if(currentSheetName.equals("2 LIAB") || currentSheetName.equals("3 OWN"))
        {

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                if (sheet.getSheetName().equals(currentSheetName))
                {

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.equals(""))
                                continue;

                            String currentXmlTag = xmlPattern;

                            if (currentXmlTag != null)
                            {

                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");

                                writeXml(currentXmlTag);
                            }
                        }
                    }
                }
            }

        }
        else if(currentSheetName.equals("4 IFC"))
        {
            int counter = 0;
            String reportDataCols [] = new String[] {"0060", "0070", "0080"};
            String valZ0400_0020 = "";
            String valZ0400_0040 = "";
            String valZ0400_0050 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ0400_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020
                valZ0400_0040 = getCellValue(row.getCell(stringToInt.get(startColumn) + 3));//col 0040
                valZ0400_0050 = getCellValue(row.getCell(stringToInt.get(startColumn) + 4));//col 0050

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0400_0020", valZ0400_0020);
                    currentXmlTag = currentXmlTag.replace("@Z0400_0040", valZ0400_0040);
                    currentXmlTag = currentXmlTag.replace("@Z0400_0050", getFormattedCellValue(valZ0400_0050));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valZ0400_0020 = getCellValue(row.getCell(columnIndex + 1));
                        valZ0400_0040 = getCellValue(row.getCell(columnIndex + 3));
                        valZ0400_0050 = getCellValue(row.getCell(columnIndex + 4));
                    }


                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0040") || colCode.equals("0050"))
                            continue;
                        if(Arrays.asList(reportDataCols).contains(colCode))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0400_0020", valZ0400_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0400_0040", valZ0400_0040);
                            currentXmlTag = currentXmlTag.replace("@Z0400_0050", getFormattedCellValue(valZ0400_0050));


                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0400_0020", valZ0400_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0400_0040", valZ0400_0040);
                            currentXmlTag = currentXmlTag.replace("@Z0400_0050", getFormattedCellValue(valZ0400_0050));
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("5.1 MCP 1") || currentSheetName.equals("5.2 MCP 2"))
        {
            int counter = 0;
            String valZ050X_0020 = "";
            String valZ050X_0060 = "";
            Row row = null;
            try
            {
                row = sheet.getRow(startRow);
            }
            catch(Exception ex)
            {
                System.out.println("EROARE");
            }
            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ050X_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020
                valZ050X_0060 = getCellValue(row.getCell(stringToInt.get(startColumn) + 5));//col 0060

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    if(currentSheetName.equals("5.1 MCP 1"))
                    {
                        currentXmlTag = currentXmlTag.replace("@Z0501_0020", valZ050X_0020);
                        currentXmlTag = currentXmlTag.replace("@Z0501_0060", getFormattedCellValue(valZ050X_0060));
                    }
                    else
                    {
                        currentXmlTag = currentXmlTag.replace("@M336_TOP10", valZ050X_0020);
                        currentXmlTag = currentXmlTag.replace("@Z0502_0060", getFormattedCellValue(valZ050X_0060));
                    }

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valZ050X_0020 = getCellValue(row.getCell(columnIndex + 1));
                        valZ050X_0060 = getCellValue(row.getCell(columnIndex + 5));
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0060"))
                            continue;
                        if(colCode.equals("0070") )
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            if(currentSheetName.equals("5.1 MCP 1"))
                            {
                                currentXmlTag = currentXmlTag.replace("@Z0501_0020", valZ050X_0020);
                                currentXmlTag = currentXmlTag.replace("@Z0501_0060", getFormattedCellValue(valZ050X_0060));
                            }
                            else
                            {
                                currentXmlTag = currentXmlTag.replace("@M336_TOP10", valZ050X_0020);
                                currentXmlTag = currentXmlTag.replace("@Z0502_0060", getFormattedCellValue(valZ050X_0060));
                            }

                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + getFormattedCellValue(cellValue) + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            if(currentSheetName.equals("5.1 MCP 1"))
                            {
                                currentXmlTag = currentXmlTag.replace("@Z0501_0020", valZ050X_0020);
                                currentXmlTag = currentXmlTag.replace("@Z0501_0060", getFormattedCellValue(valZ050X_0060));
                            }
                            else
                            {
                                currentXmlTag = currentXmlTag.replace("@M336_TOP10", valZ050X_0020);
                                currentXmlTag = currentXmlTag.replace("@Z0502_0060", getFormattedCellValue(valZ050X_0060));
                            }
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("6 DIS"))
        {
            int counter = 0;
            String valZ0600_0020 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ0600_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0600_0020", valZ0600_0020);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                        valZ0600_0020 = getCellValue(row.getCell(columnIndex + 1));

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020"))
                            continue;
                        if(colCode.equals("0040") || colCode.equals("0060"))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0600_0020", valZ0600_0020);

                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + getFormattedCellValue(cellValue) + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0600_0020", valZ0600_0020);
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("7.1 FUNC 1"))
        {
            boolean writeDummyRow ;

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfB272 = "";
                writeDummyRow = true;

                if (sheet.getSheetName().startsWith(currentSheetName))
                {
                    valueOfB272 = getCellValue(sheet, tableB272Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {
                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.equals(""))
                                continue;

                            String currentXmlTag = xmlPattern;

							/*
							if (currentXmlTag != null && writeDummyRow == true)
							{

								currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
								currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
								currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
								currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
								currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
								currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
								currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
								currentXmlTag = currentXmlTag.replace("@B272", valueOfB272);

								writeDummyRow = false;

								writeXml(currentXmlTag);
							}
							*/

                            if (currentXmlTag != null)
                            {
                                if(colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0040"))
                                {
                                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                    currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                                    currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                    currentXmlTag = currentXmlTag.replace("@B272", valueOfB272);
                                }
                                else
                                {
                                    currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                    currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                                    currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode/*"d.x"*/);
                                    currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                    currentXmlTag = currentXmlTag.replace("@B272", valueOfB272);
                                }

                                writeXml(currentXmlTag);
                            }
                        }
                    }
                }
            }
        }
        else if(currentSheetName.equals("7.2 FUNC 2"))
        {
            int counter = 0;
            String valZ0100_0020 = "";
            String valZ0702_0020 = "";
            String valB272 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valB272 = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valZ0702_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020
                valZ0100_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 3));//col 0040

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0100_0020", valZ0100_0020);
                    currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                    currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valB272 = getCellValue(row.getCell(columnIndex));//col 0010
                        valZ0702_0020 = getCellValue(row.getCell(columnIndex + 1));//col 0020
                        valZ0100_0020 = getCellValue(row.getCell(columnIndex + 3));//col 0040
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0010") || colCode.equals("0020") || colCode.equals("0040"))
                            continue;
                        if(colCode.equals("0050"))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0100_0020", valZ0100_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));

                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0100_0020", valZ0100_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        /**********************************************************DE IMPLEMENTAT***************************************************************/
        else if(currentSheetName.equals("7.3 FUNC 3"))
        {
            int counter = 0;
            String valZ0100_0020 = "";
            String valZ0703_0020 = "";
            //String valB272 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                //valB272 = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valZ0703_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020
                valZ0100_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 4));//col 0050

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0100_0020", valZ0100_0020);
                    currentXmlTag = currentXmlTag.replace("@Z0703_0020", valZ0703_0020);
                    currentXmlTag = currentXmlTag.replace("B272=\"@B272\"", "");

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        //valB272 = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                        valZ0703_0020 = getCellValue(row.getCell(columnIndex + 1));//col 0020
                        valZ0100_0020 = getCellValue(row.getCell(columnIndex + 4));//col 0050

                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0050"))
                            continue;
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0100_0020", valZ0100_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0703_0020", valZ0703_0020);
                            currentXmlTag = currentXmlTag.replace("B272=\"@B272\"", "");

                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        /**********************************************************DE IMPLEMENTAT***************************************************************/
        else if(currentSheetName.equals("7.4 FUNC 4"))
        {
            int counter = 0;
            String valZ0702_0020 = "";
            String valZ0703_0020 = "";
            String valB272 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valB272 = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valZ0702_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020
                valZ0703_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 3));//col 0040

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0703_0020", valZ0703_0020);
                    currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                    currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valB272 = getCellValue(row.getCell(columnIndex));//col 0010
                        valZ0702_0020 = getCellValue(row.getCell(columnIndex + 1));//col 0020
                        valZ0703_0020 = getCellValue(row.getCell(columnIndex + 3));//col 0040
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0010") || colCode.equals("0020") || colCode.equals("0040"))
                            continue;
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0703_0020", valZ0703_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("8 SERV"))
        {
            int counter = 0;
            String valZ0702_0020 = "";
            String valZ0800_0010 = "";
            String valZ0800_0030 = "";
            String valZ0800_0050 = "";
            String valB272 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ0800_0010 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0010, incepe cu 0005
                valZ0800_0030 = getCellValue(row.getCell(stringToInt.get(startColumn) + 3));//col 0030
                valZ0800_0050 = getCellValue(row.getCell(stringToInt.get(startColumn) + 5));//col 0050
                valB272 = getCellValue(row.getCell(stringToInt.get(startColumn) + 7));//col 0070
                valZ0702_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 8));//col 0080

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0800_0010", getFormattedCellValue(valZ0800_0010));
                    currentXmlTag = currentXmlTag.replace("@Z0800_0030", valZ0800_0030);
                    currentXmlTag = currentXmlTag.replace("@Z0800_0050", valZ0800_0050);
                    currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                    currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0005"))
                    {
                        valZ0800_0010 = getCellValue(row.getCell(columnIndex + 1));//col 0010, incepe cu 0005
                        valZ0800_0030 = getCellValue(row.getCell(columnIndex + 3));//col 0030
                        valZ0800_0050 = getCellValue(row.getCell(columnIndex + 5));//col 0050
                        valB272 = getCellValue(row.getCell(columnIndex + 7));//col 0070
                        valZ0702_0020 = getCellValue(row.getCell(columnIndex + 8));//col 0080

                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0010") || colCode.equals("0030") || colCode.equals("0050") || colCode.equals("0070") || colCode.equals("0080"))
                            continue;
                        if(colCode.equals("0005"))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0800_0010", getFormattedCellValue(valZ0800_0010));
                            currentXmlTag = currentXmlTag.replace("@Z0800_0030", valZ0800_0030);
                            currentXmlTag = currentXmlTag.replace("@Z0800_0050", valZ0800_0050);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));

                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + getFormattedCellValue(cellValue) + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0800_0010", getFormattedCellValue(valZ0800_0010));
                            currentXmlTag = currentXmlTag.replace("@Z0800_0030", valZ0800_0030);
                            currentXmlTag = currentXmlTag.replace("@Z0800_0050", valZ0800_0050);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("9 FMI"))
        {
            int counter = 0;
            String valZ0702_0020 = "";
            String valZ0900_0020 = "";
            String valZ0900_0070 = "";
            String valB272 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ0900_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020
                valB272 = getCellValue(row.getCell(stringToInt.get(startColumn) + 2));//col 0030
                valZ0702_0020 = getCellValue(row.getCell(stringToInt.get(startColumn) + 3));//col 0040
                valZ0900_0070 = getCellValue(row.getCell(stringToInt.get(startColumn) + 6));//col 0070

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z0900_0020", valZ0900_0020);
                    currentXmlTag = currentXmlTag.replace("@Z0900_0070", valZ0900_0070);
                    currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                    currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valZ0900_0020 = getCellValue(row.getCell(columnIndex + 1));//col 0020
                        valB272 = getCellValue(row.getCell(columnIndex + 2));//col 0030
                        valZ0702_0020 = getCellValue(row.getCell(columnIndex + 3));//col 0040
                        valZ0900_0070 = getCellValue(row.getCell(columnIndex + 6));//col 0070

                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0040") || colCode.equals("0070"))
                            continue;
                        if(colCode.equals("0005"))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0900_0020", valZ0900_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0900_0070", valZ0900_0070);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z0900_0020", valZ0900_0020);
                            currentXmlTag = currentXmlTag.replace("@Z0900_0070", valZ0900_0070);
                            currentXmlTag = currentXmlTag.replace("@Z0702_0020", getFormattedCellValue(valZ0702_0020));
                            currentXmlTag = currentXmlTag.replace("@B272", getFormattedCellValue(valB272));
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("10.1 CIS 1"))
        {
            int counter = 0;
            String valZ1001_0010 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ1001_0010 = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z1001_0010", valZ1001_0010);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valZ1001_0010 = getCellValue(row.getCell(columnIndex));//col 0010

                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0010"))
                            continue;
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z1001_0010", valZ1001_0010);
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("10.2 CIS 2"))
        {
            int counter = 0;
            String valZ1001_0010 = "";
            String valZ1002_0030 = "";
            String valZ1002_0040 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valZ1001_0010 = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valZ1002_0030 = getCellValue(row.getCell(stringToInt.get(startColumn) + 2));//col 0030
                valZ1002_0040 = getCellValue(row.getCell(stringToInt.get(startColumn) + 3));//col 0040

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@Z1001_0010", valZ1001_0010);
                    currentXmlTag = currentXmlTag.replace("@Z1002_0030", valZ1002_0030);
                    currentXmlTag = currentXmlTag.replace("@Z1002_0040", valZ1002_0040);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        valZ1001_0010 = getCellValue(row.getCell(columnIndex));//col 0010
                        valZ1002_0030 = getCellValue(row.getCell(columnIndex + 2));//col 0030
                        valZ1002_0040 = getCellValue(row.getCell(columnIndex + 3));//col 0040
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0010") || colCode.equals("0030") || colCode.equals("0040"))
                            continue;
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata","allocated_mrel_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", "d.x");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + getFormattedCellValue(cellValue) + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@Z1001_0010", valZ1001_0010);
                            currentXmlTag = currentXmlTag.replace("@Z1002_0030", valZ1002_0030);
                            currentXmlTag = currentXmlTag.replace("@Z1002_0040", valZ1002_0040);
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }//Final

    }

    //sheet-uri default Finrep Abacus
    public static void writeXmlFinrepDefault()
    {
        if(currentSheetName.startsWith("20."))
        {
            //caut sheet-urile dinamice
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfC010 = "";

                if (sheet.getSheetName().startsWith(currentSheetName))
                {

                    valueOfC010 = abacusCountryCodes.get(getCellValue(sheet, tableC010Cell));

                    if(valueOfC010 == null || valueOfC010.equals(""))
                        valueOfC010 = getCellValue(sheet, tableC010Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.trim().equals(""))
                                continue;

                            String currentXmlTag = xmlPattern;

                            if (currentXmlTag != null)
                            {
                                currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                currentXmlTag = currentXmlTag.replace("@C007", "RON");
                                currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                                currentXmlTag = currentXmlTag.replace("@TK_FLAG", "0");
                                currentXmlTag = currentXmlTag.replace("@C010", valueOfC010);
                                currentXmlTag = currentXmlTag.replace("@C011", valueOfC010);
                                //C010="@C010" C011="@C011" TK_FLAG="@TK_FLAG"
                                writeXml(currentXmlTag);
                            }
                        }
                    }
                }
            }
        }

        else if(!currentSheetName.equals("40.01") && !currentSheetName.equals("40.02"))
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {
                        currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@rowname", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@C007", "RON");
                        currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                        currentXmlTag = currentXmlTag.replace("@TK_FLAG", "0");
                        currentXmlTag = currentXmlTag.replace(" C010=\"@C010\"", "");
                        currentXmlTag = currentXmlTag.replace(" C011=\"@C011\"", "");
                        //C010="@C010" C011="@C011" TK_FLAG="@TK_FLAG"
                        writeXml(currentXmlTag);
                    }
                }
            }
        }
        else if(currentSheetName.equals("40.01") || currentSheetName.equals("40.02"))
        {
            int counter = 0;
            String investeeValue = "";
            String c231Value = "";
            String groupfieldValue = "";

            String [] reportDataCols = new String[]{"050","060","070","080","110","120","160","170","180","190"};
            String [] allocatedTextCols = new String[]{"010","030","090","100"};
            String [] allocatedIntCols = new String[]{"095","130","140","150"};

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;
                investeeValue = "";
                groupfieldValue = "";
                c231Value = "";

                String currentXmlTag = xmlPattern;

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(currentSheetName.equals("40.01") && colCode.equals("010"))
                        investeeValue = getCellValue(row.getCell(columnIndex + 1));
                    if(currentSheetName.equals("40.02") && colCode.equals("010"))
                    {
                        investeeValue = getCellValue(row.getCell(columnIndex + 1));
                        groupfieldValue = getCellValue(row.getCell(columnIndex + 3));
                        c231Value = cellValue;
                    }


                    if (currentXmlTag != null)
                    {
                        if(currentSheetName.equals("40.01") && colCode.equals("020"))
                            continue;
                        if((currentSheetName.equals("40.01") && Arrays.asList(reportDataCols).contains(colCode)) ||
                                (currentSheetName.equals("40.02") && (colCode.equals("060") || colCode.equals("070") || colCode.equals("080"))))
                        {
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.N");
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@C007", "RON");
                            currentXmlTag = currentXmlTag.replace("@J073", "EUR");
                            currentXmlTag = currentXmlTag.replace("@TK_FLAG", "0");
                            currentXmlTag = currentXmlTag.replace("@C231", c231Value);
                            currentXmlTag = currentXmlTag.replace("@GROUPFIELD", groupfieldValue);
                            currentXmlTag = currentXmlTag.replace(" C010=\"@C010\"", "");
                            currentXmlTag = currentXmlTag.replace(" C011=\"@C011\"", "");

                            if(!investeeValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INVESTEE", investeeValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INVESTEE=\"@INVESTEE\"", "");

                        }
                        else if((currentSheetName.equals("40.01") && Arrays.asList(allocatedTextCols).contains(colCode)) ||
                                (currentSheetName.equals("40.02") && (colCode.equals("030") || colCode.equals("050") || colCode.equals("020"))))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_finrep_text");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.N");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@TK_FLAG", "0");
                            currentXmlTag = currentXmlTag.replace("@C231", c231Value);
                            currentXmlTag = currentXmlTag.replace("@GROUPFIELD", groupfieldValue);
                            currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                            currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                            currentXmlTag = currentXmlTag.replace(" C007=\"@C007\"", "");
                            currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                            currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                            currentXmlTag = currentXmlTag.replace(" C010=\"@C010\"", "");
                            currentXmlTag = currentXmlTag.replace(" C011=\"@C011\"", "");

                            if(!investeeValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INVESTEE", investeeValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INVESTEE=\"@INVESTEE\"", "");
                        }
                        else if(currentSheetName.equals("40.01") && Arrays.asList(allocatedIntCols).contains(colCode))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_finrep_int");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.N");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "VALUE=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@TK_FLAG", "0");
                            currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                            currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                            currentXmlTag = currentXmlTag.replace(" C007=\"@C007\"", "");
                            currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                            currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                            currentXmlTag = currentXmlTag.replace(" C010=\"@C010\"", "");
                            currentXmlTag = currentXmlTag.replace(" C011=\"@C011\"", "");

                            if(!investeeValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INVESTEE", investeeValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INVESTEE=\"@INVESTEE\"", "");
                        }
                        else if(currentSheetName.equals("40.01") && colCode.equals("040"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_finrep_date");
                            currentXmlTag = currentXmlTag.replace("@f001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@colname", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@rowname", "d.N");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "DATEVAL=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@TK_FLAG", "0");
                            currentXmlTag = currentXmlTag.replace(" NULLCHECK=\"@NULLCHECK\"", "");
                            currentXmlTag = currentXmlTag.replace(" TYPE=\"@TYPE\"", "");
                            currentXmlTag = currentXmlTag.replace(" C007=\"@C007\"", "");
                            currentXmlTag = currentXmlTag.replace(" J073=\"@J073\"", "");
                            currentXmlTag = currentXmlTag.replace(" P024=\"0\"", "");
                            currentXmlTag = currentXmlTag.replace(" C010=\"@C010\"", "");
                            currentXmlTag = currentXmlTag.replace(" C011=\"@C011\"", "");

                            if(!investeeValue.equals(""))
                                currentXmlTag = currentXmlTag.replace("@INVESTEE", investeeValue);
                            else
                                currentXmlTag = currentXmlTag.replace(" INVESTEE=\"@INVESTEE\"", "");
                        }
                        else
                            continue;

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
    }


    public static boolean sheetHasData(Sheet sheet)
    {
        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);
            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);
                String cellValue = getCellValue(cell);
                if(cellValue != null && !cellValue.trim().equals(""))
                    return true;
            }
        }
        return false;
    }

    //sheet-uri default Finrep bnr
    public static void writeXmlFinrepBnr()
    {
        int valueCounter = 1;


        if(getParamsFinrepBnr == true)
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                Sheet paramSheet = workbook.getSheetAt(sheetIndex);

                if (paramSheet.getSheetName().equals("Parametrii"))
                {
                    senderIDBnr = getCellValue(paramSheet,"C3");
                    sendingDateBnr = getCellValue(paramSheet,"C4");
                    messageTypeBnr = getCellValue(paramSheet,"C5");
                    refDateBnr = getCellValue(paramSheet,"C6");
                    senderMessageIdBnr = getCellValue(paramSheet,"C7");
                    operationTypeBnr = getCellValue(paramSheet,"C8");
                    corelatedWithBnr = getCellValue(paramSheet,"C9");
                    correctionOfBnr = getCellValue(paramSheet,"C10");

                    getParamsFinrepBnr = false;

                    writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
                    if(messageTypeBnr.equalsIgnoreCase("RFC415-06"))
                    {
                        writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro RFC415-06.xsd\">");
                    }
                    else	writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro RFC400-05.xsd\">");
                    writeXml("	<Header>");
                    writeXml("		<SenderId>" + senderIDBnr + "</SenderId>");
                    writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(DateUtil.getJavaDate(Double.parseDouble(sendingDateBnr))) + "</SendingDate>");
                    writeXml("		<MessageType>" + messageTypeBnr + "</MessageType>");
                    writeXml("		<RefDate>" + new SimpleDateFormat("yyyy-MM-dd").format(DateUtil.getJavaDate(Double.parseDouble(refDateBnr))) + "</RefDate>");
                    writeXml("		<SenderMessageId>" + senderMessageIdBnr + "</SenderMessageId>");
                    writeXml("		<OperationType>" + operationTypeBnr + "</OperationType>");
                    //writeXml("		<CodeLEI>" + LEICode + "</CodeLEI>");
                    writeXml("	</Header>");
                    writeXml("	<Body>");

                    break;
                }
            }

        if (messageTypeBnr.equalsIgnoreCase("RFC415-06") &&
                !(currentSheetName.startsWith("F41") ||
                        currentSheetName.startsWith("F42") ||
                        currentSheetName.startsWith("F44") ||
                        currentSheetName.startsWith("F45") ||
                        currentSheetName.startsWith("F46") ||
                        currentSheetName.startsWith("F47") )) {
            return;
        }
        if (messageTypeBnr.equalsIgnoreCase("RFC400-05") &&
                (currentSheetName.startsWith("F41") ||
                        currentSheetName.startsWith("F42") ||
                        currentSheetName.startsWith("F44") ||
                        currentSheetName.startsWith("F45") ||
                        currentSheetName.startsWith("F46") ||
                        currentSheetName.startsWith("F47"))) {
            return;
        }



        if(currentSheetName.startsWith("F20.") && (reportingEntity.substring(0, 3).equals("302") || reportingEntity.substring(0, 3).equals("327") || reportingEntity.substring(0, 3).equals("319")|| reportingEntity.substring(0, 3).equals("349")))
        {
            int skipValue = 0;
            int tablePosition = 0;
            int rowCodeC010Cell;
            String finaltableC010Cell = "";

            if(currentSheetName.equals("F20.04"))
            {
                try
                {
                    if(DateUtil.getJavaDate(Double.parseDouble(refDateBnr)).before(refDateBnrFormat.parse("30/06/2021")))
                        skipValue = 31;
                    else
                        skipValue = 32;
                }
                catch (ParseException e)
                {
                    skipValue = 32;
                }
            }
            else if(currentSheetName.equals("F20.05"))
                skipValue = 10;
            else if(currentSheetName.equals("F20.06"))
                skipValue = 18;
            else if(currentSheetName.equals("F20.07"))
                skipValue = 26;

            boolean reportApendixIfdata = true;

            valueOfC010 = "";

            valueOfC010 = getCellValue(sheet, tableC010Cell).trim();



            if(sheetHasData(sheet) || 1 == 1)
            {
                while(valueOfC010 != null && !valueOfC010.trim().equals(""))
                {
                    reportApendixIfdata = true;

                    for(int rowIndex = startRow + tablePosition * skipValue; rowIndex <= endRow + tablePosition * skipValue; rowIndex++)
                    {
                        valueCounter = 0;
                        boolean newRow = true;

                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }
                        if(rowCode.equals("") || rowCode == null)
                            continue;
                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            valueCounter++;
                            //System.out.println(rowIndex + "       " + columnIndex);
                            Cell cell = row.getCell(columnIndex);

                            String cellValue;

                            if(cell == null || isIgnoredColor(cell))
                                cellValue = "0";
                            else
                                cellValue = getCellValue(cell);
                            if(cellValue.equals(""))
                                cellValue = "0";

                            if(reportApendixIfdata == true)
                            {
                                writeXml("		<Appendix>");
                                writeXml("			<AppendixCode>" + reportName + "</AppendixCode>");
                                writeXml("			<Country>" + valueOfC010 + "</Country>");
                                writeXml("			<Items>");
                                reportApendixIfdata =  false;
                            }
                            if(newRow == true)
                            {
                                writeXml("				<Item>");
                                writeXml("					<Code>" + rowCode + "</Code>");
                                newRow = false;
                            }
                            writeXml("					<Value"+ valueCounter + ">" + cellValue + "</Value"+ valueCounter + ">");
                        }

                        if(newRow == false)
                            writeXml("				</Item>");

                    }
                    if(reportApendixIfdata == false)
                    {
                        writeXml("			</Items>");
                        writeXml("		</Appendix>");
                    }

                    tablePosition++;

                    rowCodeC010Cell = Integer.parseInt(tableC010Cell.replaceAll("[a-zA-Z]+", ""));
                    rowCodeC010Cell = skipValue * tablePosition + rowCodeC010Cell;

                    finaltableC010Cell = tableC010Cell.replaceAll("[0-9]+", "") + rowCodeC010Cell;

                    valueOfC010 = getCellValue(sheet, finaltableC010Cell).trim();

                }

            }

        }
        else if(currentSheetName.startsWith("F20."))
        {
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                boolean reportApendixIfdata = true;

                sheet = workbook.getSheetAt(sheetIndex);

                if (sheet.getSheetName().startsWith(currentSheetName))
                {
                    if(!sheetHasData(sheet))
                        continue;

                    valueOfC010 = getCellValue(sheet, tableC010Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        valueCounter = 0;
                        boolean newRow = true;

                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }
                        if(rowCode.equals("") || rowCode == null)
                            continue;
                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            valueCounter++;
                            //System.out.println(rowIndex + "       " + columnIndex);
                            Cell cell = row.getCell(columnIndex);

                            String cellValue;

                            if(cell == null || isIgnoredColor(cell))
                                cellValue = "0";
                            else
                                cellValue = getCellValue(cell);
                            if(cellValue.equals(""))
                                cellValue = "0";

                            if(reportApendixIfdata == true)
                            {
                                writeXml("		<Appendix>");
                                writeXml("			<AppendixCode>" + reportName + "</AppendixCode>");
                                writeXml("			<Country>" + valueOfC010 + "</Country>");
                                writeXml("			<Items>");
                                reportApendixIfdata =  false;
                            }
                            if(newRow == true)
                            {
                                writeXml("				<Item>");
                                writeXml("					<Code>" + rowCode + "</Code>");
                                newRow = false;
                            }
                            writeXml("					<Value"+ valueCounter + ">" + cellValue + "</Value"+ valueCounter + ">");
                        }

                        if(newRow == false)
                            writeXml("				</Item>");

                    }
                    if(reportApendixIfdata == false)
                    {
                        writeXml("			</Items>");
                        writeXml("		</Appendix>");
                    }

                }
            }
        }
        else
        {
            boolean reportApendixIfdata = true;

            if(sheetHasData(sheet))

                for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    valueCounter = 0;
                    boolean newRow = true;

                    Row row = sheet.getRow(rowIndex);

                    String rowCode, xmlRowCode = "";

                    if (tableRowCodeColumn.equals("*"))
                    {
                        rowCode = "010";
                        xmlRowCode = "010";
                    }
                    else
                    {
                        if(xmlRowNamePattern.equals("ABACUS"))
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            if(abacusRowNames.containsKey(rowCode))
                                xmlRowCode = abacusRowNames.get(rowCode);
                        }
                        else
                        {
                            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                            xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                        }
                    }

                    if(rowCode.equals("") || rowCode == null)
                        continue;
                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        valueCounter++;
                        //System.out.println(rowIndex + "       " + columnIndex);
                        Cell cell = row.getCell(columnIndex);
                        String cellValue;
					/*if(isIgnoredColor(cell))
						continue;

					String cellValue = getCellValue(cell);

					if (cellValue.trim().equals(""))
						continue;*/

                        if(cell == null || isIgnoredColor(cell))
                            cellValue = "0";
                        else
                            cellValue = getCellValue(cell);
                        if(cellValue.equals(""))
                            cellValue = "0";

                        if(reportApendixIfdata == true)
                        {
                            writeXml("		<Appendix>");
                            writeXml("			<AppendixCode>" + reportName + "</AppendixCode>");
                            writeXml("			<Items>");
                            reportApendixIfdata =  false;
                        }
                        if(newRow == true)
                        {
                            writeXml("				<Item>");
                            writeXml("					<Code>" + rowCode + "</Code>");
                            newRow = false;
                        }
                        writeXml("					<Value"+ valueCounter + ">" + cellValue + "</Value"+ valueCounter + ">");
                    }

                    if(newRow == false)
                        writeXml("				</Item>");

                }
            if(reportApendixIfdata == false)
            {
                writeXml("			</Items>");
                writeXml("		</Appendix>");
            }
        }

    }


    public static void writeXmlCovid19()
    {
        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);

            String rowCode, xmlRowCode = "";

            if (tableRowCodeColumn.equals("*"))
            {
                rowCode = "010";
                xmlRowCode = "010";
            }
            else
            {
                if(xmlRowNamePattern.equals("ABACUS"))
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                    if(abacusRowNames.containsKey(rowCode))
                        xmlRowCode = abacusRowNames.get(rowCode);
                }
                else
                {
                    rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                    xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                }
            }

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                String colCode, xmlColCode = "";

                if (tableColumnCodeRow.equals("*"))
                {
                    colCode = "010";
                    xmlColCode = "010";
                }
                else
                {

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                    xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                }

                String cellValue = getCellValue(cell);

                if (cellValue.trim().equals(""))
                    continue;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    if((currentSheetName.equals("F 93.01") || currentSheetName.equals("F 93.02.b")) && colCode.equals("0030"))
                        currentXmlTag = currentXmlTag.replace("allocated_covid19_its", "allocated_covid19_text");
                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@REPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                    if((currentSheetName.equals("F 93.01") || currentSheetName.equals("F 93.02.b")) && colCode.equals("0030"))
                        currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                    else
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue.replace(",", ""));
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");


                    writeXml(currentXmlTag);
                }
            }
        }
    }


    //sheet-uri default Funding Plan
    public static void writeXmlFundingPlan()
    {
        if(currentSheetName.startsWith("2C"))
        {
            //caut sheet-urile dinamice
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfC007 = "";

                if (sheet.getSheetName().startsWith("2C"))
                {

                    valueOfC007 = abacusCountryCodes.get(getCellValue(sheet, tableC007Cell));

                    if(valueOfC007 == null || valueOfC007.equals(""))
                        valueOfC007 = getCellValue(sheet, tableC007Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.trim().equals(""))
                                continue;

                            String currentXmlTag = xmlPattern;

                            if (currentXmlTag != null)
                            {
                                currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                                currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                currentXmlTag = currentXmlTag.replace("@C007", valueOfC007);
                                writeXml(currentXmlTag);
                            }
                        }
                    }
                }
            }
        }
        else
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {
                        if(currentSheetName.startsWith("2A") && reportName.equals("fp_p02_03") && colCode.equals("060"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_fp_text");
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT0=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"","REPORTNAME=\"" + reportName + "\"");
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        }
                        currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        writeXml(currentXmlTag);
                    }
                }
            }
        }
    }


    public static void writeXmlSbp()
    {

        if(currentSheetName.equals("101"))
        {
            int counter = 0;
            String valInc = "";
            String valSbp094 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valInc = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010

                valSbp094 = valInc.replaceAll(".*?_(C[A-Z])_.*|.*", "$1");


                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@INC", valInc);
                    if (valSbp094 != null)
                        currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                    else
                        currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                    //currentXmlTag = currentXmlTag.replace("@Z0400_0050", getFormattedCellValue(valZ0400_0050));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0070"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);
                            if (valSbp094 != null)
                                currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                            else
                                currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                        }
                        else if(colCode.equals("0050"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_date");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            try
                            {
                                currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "DATEVAL=\"" + dateFromNumber(cellValue) + "\"");
                            }
                            catch(Exception ex)
                            {
                                currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "DATEVAL=\"" + cellValue + "\"");
                            }
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);
                            if (valSbp094 != null)
                                currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                            else
                                currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);
                            if (valSbp094 != null)
                                currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                            else
                                currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("102") || currentSheetName.equals("103"))
        {
            int counter = 0;
            String valPbe = "";
            String valSbp094 = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010

                valSbp094 = valPbe.replaceAll(".*?_(C[A-Z])_.*|.*", "$1");


                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                    if (valSbp094 != null)
                        currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                    else
                        currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                    //currentXmlTag = currentXmlTag.replace("@Z0400_0050", getFormattedCellValue(valZ0400_0050));

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0070"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            if (valSbp094 != null)
                                currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                            else
                                currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            if (valSbp094 != null)
                                currentXmlTag = currentXmlTag.replace("@SBP094", valSbp094);
                            else
                                currentXmlTag = currentXmlTag.replace("SBP094=\"@SBP094\"", "");
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("105.01"))
        {
            int counter = 0;
            String valImi = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valImi = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@IMI", valImi);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0120"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@IMI", valImi);
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            if(colCode.equals("0110"))
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                            else
                                currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@IMI", valImi);
                            if(colCode.equals("0110"))
                                currentXmlTag = currentXmlTag.replace("/>", " INTVAL=\"" + cellValue + "\"/>");
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("105.02"))
        {
            int counter = 0;
            String valImi = "";
            String valPbe = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valImi = getCellValue(row.getCell(stringToInt.get(startColumn) + 1)); //col 0020
                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn))); //col 0010

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@IMI", valImi);
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0010") || colCode.equals("0020"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@IMI", valImi);
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@IMI", valImi);
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("105.03"))
        {
            int counter = 0;
            String valIrn = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valIrn= getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0005

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@IRN", valIrn);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0005"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {
                        currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                        currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                        currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                        currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@IRN", valIrn);


                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("106"))
        {
            int counter = 0;
            String valPbe = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn))); //col 0010

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);

                    writeXmlIMV(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0050") || colCode.equals("0060"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                        }
                        else if(colCode.equals("0070"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            currentXmlTag = currentXmlTag.replace("/>", " INTVAL=\"" + cellValue + "\"/>");

                        }

                        writeXmlIMV(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("107.01"))
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    String currentXmlTag = xmlPattern;
                    //xmlRowCode
                    if (currentXmlTag != null)
                    {
                        if((rowCode.equals("0050") || rowCode.equals("0060") || rowCode.equals("0090")) && colCode.equals("0010"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");

                        }


                        writeXmlRM(currentXmlTag);
                    }
                }
            }
        }
        else if (currentSheetName.startsWith("107.02"))
        {
            int counter = 0;
            String valPbe = "";
            String valCcy = "";
            String valNed = "";
            String valNedFormatted = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn)));
                valCcy = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));
                valNed = getCellValue(row.getCell(stringToInt.get(startColumn) + 2));

                try
                {
                    valNedFormatted = dateFromNumber(valNed);
                }
                catch(Exception ex)
                {
                    valNedFormatted = valNed;
                }

                if(valNedFormatted == null)
                    valNedFormatted = valNed;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                    currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                    if(valCcy != null)
                        currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                    else
                        currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                    writeXmlRM(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020") || colCode.equals("0030") || colCode.equals("0040"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue );
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                            if(valCcy != null)
                                currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                            else
                                currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                            writeXmlRM(currentXmlTag);
                        }
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if (currentSheetName.startsWith("108"))
        {
            int counter = 0;
            String valPbe = "";
            String valCcy = "";
            String valNed = "";
            String valNedFormatted = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn)));
                valCcy = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));
                valNed = getCellValue(row.getCell(stringToInt.get(startColumn) + 2));

                try
                {
                    valNedFormatted = dateFromNumber(valNed);
                }
                catch(Exception ex)
                {
                    valNedFormatted = valNed;
                }

                if(valNedFormatted == null)
                    valNedFormatted = valNed;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                    currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                    if(valCcy != null)
                        currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                    else
                        currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                    writeXmlRM(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                            if(valCcy != null)
                                currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                            else
                                currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                            writeXmlRM(currentXmlTag);
                        }
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("109.01"))
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {

                        currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                        currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                        currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");


                        writeXmlRM(currentXmlTag);
                    }
                }
            }
        }
        else if(currentSheetName.equals("109.02"))
        {
            //caut sheet-urile dinamice
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfC007 = "";

                if (sheet.getSheetName().startsWith("109.02"))
                {

                    valueOfC007 = getCellValue(sheet, tableC007Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

						/*String currentXmlTag = xmlPattern;

						if (currentXmlTag != null)
						{
							currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
							currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
							currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
							currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
							currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
							currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
							currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
							currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
							currentXmlTag = currentXmlTag.replace("@PBE", valueOfC007);


							writeXml(currentXmlTag);
						}*/

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.trim().equals(""))
                                continue;

                            String currentXmlTag = xmlPattern;

                            if (currentXmlTag != null)
                            {
                                currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                                currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                                currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                currentXmlTag = currentXmlTag.replace("@PBE", valueOfC007);

                                writeXmlRM(currentXmlTag);
                            }
                        }
                    }
                }
            }
        }
        else if (currentSheetName.startsWith("109.03"))
        {
            int counter = 0;
            String valPbe = "";
            String valCcy = "";
            String valNed = "";
            String valNedFormatted = "";

            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn)));
                valCcy = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));
                valNed = getCellValue(row.getCell(stringToInt.get(startColumn) + 2));

                try
                {
                    valNedFormatted = dateFromNumber(valNed);
                }
                catch(Exception ex)
                {
                    valNedFormatted = valNed;
                }

                if(valNedFormatted == null)
                    valNedFormatted = valNed;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                    currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                    if(valCcy != null)
                        currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                    else
                        currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                    writeXmlRM(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0020"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue );
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                            if(valCcy != null)
                                currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                            else
                                currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                            writeXmlRM(currentXmlTag);
                        }
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("110.01"))
        {
            for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                String rowCode, xmlRowCode = "";

                if (tableRowCodeColumn.equals("*"))
                {
                    rowCode = "010";
                    xmlRowCode = "010";
                }
                else
                {
                    if(xmlRowNamePattern.equals("ABACUS"))
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        if(abacusRowNames.containsKey(rowCode))
                            xmlRowCode = abacusRowNames.get(rowCode);
                    }
                    else
                    {
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                        xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                    }
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {

                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    String currentXmlTag = xmlPattern;

                    if (currentXmlTag != null)
                    {

                        currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                        currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                        currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                        currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");


                        writeXmlRM(currentXmlTag);
                    }
                }
            }
        }
        else if(currentSheetName.equals("110.02"))
        {
            //caut sheet-urile dinamice
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfC007 = "";

                if (sheet.getSheetName().startsWith("110.02"))
                {

                    valueOfC007 = getCellValue(sheet, tableC007Cell);

                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        Row row = sheet.getRow(rowIndex);

						/*String currentXmlTag = xmlPattern;

						if (currentXmlTag != null)
						{
							currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
							currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
							currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
							currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
							currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
							currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
							currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
							currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
							currentXmlTag = currentXmlTag.replace("@PBE", valueOfC007);


							writeXml(currentXmlTag);
						}*/

                        String rowCode, xmlRowCode = "";

                        if (tableRowCodeColumn.equals("*"))
                        {
                            rowCode = "010";
                            xmlRowCode = "010";
                        }
                        else
                        {
                            if(xmlRowNamePattern.equals("ABACUS"))
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                if(abacusRowNames.containsKey(rowCode))
                                    xmlRowCode = abacusRowNames.get(rowCode);
                            }
                            else
                            {
                                rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();
                                xmlRowCode = xmlRowNamePattern.replace("ITS", rowCode);
                            }
                        }

                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            Cell cell = row.getCell(columnIndex);

                            if(isIgnoredColor(cell))
                                continue;

                            String colCode, xmlColCode = "";

                            if (tableColumnCodeRow.equals("*"))
                            {
                                colCode = "010";
                                xmlColCode = "010";
                            }
                            else
                            {

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                                xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                            }

                            String cellValue = getCellValue(cell);

                            if (cellValue.trim().equals(""))
                                continue;

                            String currentXmlTag = xmlPattern;

                            if (currentXmlTag != null)
                            {
                                currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                                currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                                currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                                currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                                currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowCode);
                                currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                                currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                                currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                                currentXmlTag = currentXmlTag.replace("@PBE", valueOfC007);

                                writeXmlRM(currentXmlTag);
                            }
                        }
                    }
                }
            }
        }
        else if (currentSheetName.startsWith("110.03"))
        {
            int counter = 0;
            String valPbe = "";
            String valCcy = "";
            String valNed = "";
            String valNedFormatted = "";
            Row row = sheet.getRow(startRow);

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valPbe = getCellValue(row.getCell(stringToInt.get(startColumn)));
                valCcy = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));
                valNed = getCellValue(row.getCell(stringToInt.get(startColumn) + 2));
                try
                {
                    valNedFormatted = dateFromNumber(valNed);
                }
                catch(Exception ex)
                {
                    valNedFormatted = valNed;
                }

                if(valNedFormatted == null)
                    valNedFormatted = valNed;

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", ".0000000000");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                    currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                    if(valCcy != null)
                        currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                    else
                        currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                    writeXmlRM(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }

                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0060"))
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue );
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@PBE", valPbe);
                            currentXmlTag = currentXmlTag.replace("@NED", valNedFormatted);
                            if(valCcy != null)
                                currentXmlTag = currentXmlTag.replace("@C007", valCcy);
                            else
                                currentXmlTag = currentXmlTag.replace("C007=\"@C007\"", "");

                            writeXmlRM(currentXmlTag);
                        }
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("111"))
        {
            int counter = 0;
            String valInc = "";
            Row row = null;
            try
            {
                row = sheet.getRow(startRow);
            }
            catch(Exception ex)
            {}

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valInc = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010


                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", "0");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@INC", valInc);


                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0030"))
                        {
                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);

                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);
                        }


                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("112"))
        {
            int counter = 0;
            String valInc = "";
            String valEsc = "";

            Row row = null;
            try
            {
                row = sheet.getRow(startRow);
            }
            catch(Exception ex)
            {}

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valInc = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valEsc = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", "0");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@INC", valInc);
                    currentXmlTag = currentXmlTag.replace("@ESC", valEsc);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010") || colCode.equals("0020"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {

                        currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@INC", valInc);
                        currentXmlTag = currentXmlTag.replace("@ESC", valEsc);


                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("113"))
        {
            int counter = 0;
            String valInc = "";
            String valFty = "";

            Row row = null;
            try
            {
                row = sheet.getRow(startRow);
            }
            catch(Exception ex)
            {}

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valInc = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valFty = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", "0");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@INC", valInc);
                    currentXmlTag = currentXmlTag.replace("@FTY", valFty);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010") || colCode.equals("0020"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {
                        if(colCode.equals("0200") || colCode.equals("0600"))
                        {

                            currentXmlTag = currentXmlTag.replace("reportdata", "allocated_sbp_text");
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("PHYSICALREPORTNAME=\"@PHYSICALREPORTNAME\"", "REPORTNAME=\"" + reportName + "\"");
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("AMOUNT=\"@AMOUNT\"", "TEXT=\"" + cellValue + "\"");
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);
                            currentXmlTag = currentXmlTag.replace("@FTY", valFty);
                        }
                        else
                        {
                            currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                            currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                            currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                            currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                            currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                            currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                            currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                            currentXmlTag = currentXmlTag.replace("@INC", valInc);
                            currentXmlTag = currentXmlTag.replace("@FTY", valFty);
                        }

                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
        else if(currentSheetName.equals("114"))
        {
            int counter = 0;
            String valEsc = "";
            String valFli = "";

            Row row = null;
            try
            {
                row = sheet.getRow(startRow);
            }
            catch(Exception ex)
            {

            }

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                valFli = getCellValue(row.getCell(stringToInt.get(startColumn)));//col 0010
                valEsc = getCellValue(row.getCell(stringToInt.get(startColumn) + 1));//col 0020

                String currentXmlTag = xmlPattern;

                if (currentXmlTag != null)
                {

                    currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                    currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                    currentXmlTag = currentXmlTag.replace("@COLNAME", "DUMMY_DYN_CELL");
                    currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                    currentXmlTag = currentXmlTag.replace("@AMOUNT", "0");
                    currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                    currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                    currentXmlTag = currentXmlTag.replace("@ESC", valEsc);
                    currentXmlTag = currentXmlTag.replace("@FLI", valFli);

                    writeXml(currentXmlTag);
                }

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    String colCode, xmlColCode = "";

                    if (tableColumnCodeRow.equals("*"))
                    {
                        colCode = "010";
                        xmlColCode = "010";
                    }
                    else
                    {
                        colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                        xmlColCode = xmlColumnNamePattern.replace("ITS", colCode);
                    }

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        continue;

                    currentXmlTag = xmlPattern;

                    if(colCode.equals("0010") || colCode.equals("0020"))
                    {
                        continue;
                    }


                    if (currentXmlTag != null)
                    {

                        currentXmlTag = currentXmlTag.replace("@F001", reportingEntity);
                        currentXmlTag = currentXmlTag.replace("@PHYSICALREPORTNAME", reportName);
                        currentXmlTag = currentXmlTag.replace("@COLNAME", xmlColCode);
                        currentXmlTag = currentXmlTag.replace("@ROWNAME", xmlRowNamePattern);
                        currentXmlTag = currentXmlTag.replace("@AMOUNT", cellValue);
                        currentXmlTag = currentXmlTag.replace("@NULLCHECK", "1");
                        currentXmlTag = currentXmlTag.replace("@TYPE", "DELIVERED");
                        currentXmlTag = currentXmlTag.replace("@ESC", valEsc);
                        currentXmlTag = currentXmlTag.replace("@FLI", valFli);


                        writeXml(currentXmlTag);
                    }
                }

                row = sheet.getRow(startRow + counter);
            }
        }
    }


    /* ---------------------- RAPOARTE BNR -------------------------- */

    public static void writeXmlSpd278()
    {
        String location;
        String code;
        String value;
        String comment;

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            if (sheetIndex == 0)
            {
                writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
                writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro SPD278-01.xsd\">");
                writeXml("	<Header>");
                writeXml("		<SenderId>" + reportingEntity + "</SenderId>");
                writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
                writeXml("		<MessageType>" + "SPD278-01" + "</MessageType>");
                writeXml("		<RefDate>" + refDate + "</RefDate>");
                writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
                writeXml("		<OperationType>" + "T" + "</OperationType>");
                writeXml("	</Header>");
                writeXml("	<Body>");
                writeXml("		<Items>");
            }

            int rowCount = sheet.getLastRowNum();
            for(int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);
                location = getCellValue(row.getCell(0));
                code = getCellValue(row.getCell(1));
                value = getCellValue(row.getCell(2));
                comment = getCellValue(row.getCell(4));

                if(value.trim() != null && !value.trim().equals(""))
                {
                    writeXml("			<Item>");

                    writeXml("				<Code>" + code  + "</Code>");
                    if(location != null && !location.equals(""))
                        writeXml("				<Location>" + location  + "</Location>");
                    DecimalFormat df = new DecimalFormat("0.##");

                    writeXml("				<Value>" + df.format(Double.parseDouble(value))  + "</Value>");
                    //writeXml("        <Type>" + getCellValue(row.getCell(1))  + "</Type>");
                    if(comment != null && !comment.equals(""))
                        writeXml("				<Comments>" + comment  + "</Comments>");

                    writeXml("			</Item>");
                }
            }
        }
        writeXml("		</Items>");

    }


    public static void writeXmlSpe273()
    {
        String location;
        String code;
        String value;
        String comment;
        DecimalFormat df = new DecimalFormat("0.##");

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            if (sheetIndex == 0)
            {
                writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
                writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro SPE273-01.xsd\">");
                writeXml("	<Header>");
                writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
                writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
                writeXml("		<MessageType>" + "SPE273-01" + "</MessageType>");
                writeXml("		<RefDate>" + refDate + "</RefDate>");
                writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
                writeXml("		<OperationType>" + "T" + "</OperationType>");
                writeXml("	</Header>");
                writeXml("	<Body>");
                writeXml("		<Items>");
            }

            int rowCount = sheet.getLastRowNum();
            for(int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                try
                {
                    location = getCellValue(row.getCell(0));
                    code = getCellValue(row.getCell(1));
                    value = getCellValue(row.getCell(2));
                    comment = getCellValue(row.getCell(4));
                }
                catch (Exception ex)
                {
                    System.out.println("");
                    break;
                }
                if(value != null && !value.equals(""))
                {

                    writeXml("			<Item>");

                    writeXml("				<Code>" + code  + "</Code>");
                    if(location != null && !location.equals(""))
                        writeXml("				<Location>" + location  + "</Location>");
                    writeXml("				<Value>" + df.format(Double.parseDouble(value))  + "</Value>");
                    //writeXml("        <Type>" + getCellValue(row.getCell(1))  + "</Type>");
                    if(comment != null && !comment.equals(""))
                        writeXml("				<Comments>" + comment  + "</Comments>");

                    writeXml("			</Item>");
                }
            }
        }
        writeXml("		</Items>");

    }

    public static void writeXmlSpd270()throws SQLException
    {
        connection.prepareStatement("delete from SPD270_values where execution_id = " + String.valueOf(exId)).execute();
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String code;
        String geo;
        String mcc;
        String value1;
        String value2;
        String comment1;
        String comment2;

        BigDecimal bigD;
        //int nos = workbook.getNumberOfSheets();

        for (int sheetIndex = 0; sheetIndex < 1; sheetIndex++)
        {

            DecimalFormat df = new DecimalFormat("0.######");
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            if (sheetIndex == 0)
            {
                writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro SPD270-01.xsd\">");
                //writeXml("	<xsi:schemaLocation=\"http://extranet.bnr.ro SPD278-01.xsd\">");
                writeXml("	<Header>");
                writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
                writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
                writeXml("		<MessageType>" + "SPD270-01" + "</MessageType>");
                writeXml("		<RefDate>" + refDate + "</RefDate>");
                writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
                writeXml("		<OperationType>" + "T" + "</OperationType>");
                writeXml("	</Header>");
                writeXml("	<Body>");
                writeXml("		<Contacts>");

                for(int contactNo = 0; contactNo < nrContacte; contactNo++)
                {
                    writeXml("			<Contact>");

                    writeXml("				<Name>" + SPD270_Contacts[contactNo][1].replaceAll("\\s+", "") + "</Name>");
                    writeXml("				<Phone>" + SPD270_Contacts[contactNo][2] + "</Phone>");
                    writeXml("				<Email>" + SPD270_Contacts[contactNo][3] + "</Email>");

                    writeXml("			</Contact>");
                }

                writeXml("		</Contacts>");
                writeXml("		<Items>");
            }

            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            for(int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);
                code = getCellValue(row.getCell(0));
                geo = getCellValue(row.getCell(2));
                mcc = getCellValue(row.getCell(3));
                value1 = getCellValue(row.getCell(4));
                value2 = getCellValue(row.getCell(5));
                comment1 = getCellValue(row.getCell(12));
                comment2 = getCellValue(row.getCell(13));

                //System.out.println(row.getCell(0));

                //System.out.println(row.getCell(2));
                //if(code.equals("COD") || code.equals("9.1") || code.equals("9.1.1.1") || code.equals("9.2") || code.equals("9.3") || code.equals("9.3.1.1") || code.equals("9.4") || code.equals("9.5"))	//nu inteleg 100% cum functioneaza asta, trebuie sa remediez schema mai tarziu
                if(isIgnoredColor(row.getCell(0)))
                    //System.out.println(getFillColorHex(row.getCell(0)));
                    continue;

                //System.out.println(row.getCell(0));

                if(code.trim() != null && !code.trim().equals("") &&
                        (geo != null && !geo.equals("") || mcc != null && !mcc.equals("") || value1.trim() != null && !value1.trim().equals("") || comment1.trim() != null && !comment1.trim().equals("") || value2.trim() != null && !value2.trim().equals("") || comment2.trim() != null && !comment2.trim().equals(""))
                )
                {

                    if(code.equals("9.3.1.1.1") || code.equals("9.3.1.1.2"))
                    {
                        if(geo.equals("BJ"))
                        {
                            int x = 1;
                        }
                        SPD270MapValue1.put(code + "#" + geo, SPD270MapValue1.get(code + "#" + geo) == null ? Double.parseDouble(value1) : SPD270MapValue1.get(code + "#" + geo) + Double.parseDouble(value1));
                        SPD270MapValue2.put(code + "#" + geo, SPD270MapValue2.get(code + "#" + geo) == null ? Double.parseDouble(value2) : SPD270MapValue2.get(code + "#" + geo) + Double.parseDouble(value2));
                    }


                    writeXml("			<Item>");

                    writeXml("				<Code>" + code  + "</Code>");
                    if(geo != null && !geo.equals(""))
                        writeXml("				<Geo>" + geo  + "</Geo>");


                    if(mcc != null && !mcc.equals(""))
                        writeXml("				<Mcc>" + mcc  + "</Mcc>");

                    if(value1.trim() != null && !value1.trim().equals(""))
                        writeXml("				<Value1>" + String.valueOf(Integer.parseInt(value1))  + "</Value1>");

                    if(comment1.trim() != null && !comment1.trim().equals(""))
                        writeXml("				<Comment1>" + comment1  + "</Comment1>");

                    if(value2.trim() != null && !value2.trim().equals(""))
                    {
                        bigD = new BigDecimal(value2);
                        writeXml("				<Value2>" + df.format(bigD)  + "</Value2>");
                    }

                    if(comment2.trim() != null && !comment2.trim().equals(""))
                        writeXml("				<Comment2>" + comment2  + "</Comment2>");


                    writeXml("			</Item>");

			    	/*
					preparedStmt.setInt(1, exId);
					preparedStmt.setString(2, code);
					preparedStmt.setString(3, geo);
					preparedStmt.setString(4, mcc);
					preparedStmt.setString(5, value1);
					preparedStmt.setString(6, value2);
					preparedStmt.setString(7, comment1);
					preparedStmt.setString(8, comment2);
					preparedStmt.addBatch();
					*/
                }


            }
        }



        writeXml("		</Items>");
        //preparedStmt.executeBatch();

    }



    public static void writeXmlSTM231() throws SQLException {

        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String xsd;
        int cnt = 0;

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        String executionDateTime = LocalDateTime.now().format(formatter);
        Instant now = Instant.now();
        String currentDate = now.toString();
        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if (valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if (tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);

        int SenderMSGIDInt = Integer.parseInt(SenderMSGID);

        if (!headerWritten) {
            writeXml("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"
                    + "<Message xmlns=\"http://extranet.bnr.ro\" xsi:schemaLocation=\"http://extranet.bnr.ro STM231-03.xsd\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\r\n"
                    + "  <Header>\r\n"
                    + "    <SenderId>" + reportingEntity.replace("i", "") + "</SenderId>\r\n"
                    + "    <SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>\r\n"
                    + "    <MessageType>STM231-03</MessageType>\r\n"
                    + "    <RefDate>" + refDate + "</RefDate>\r\n"
                    + "    <SenderMessageId>" + SenderMSGIDInt + "</SenderMessageId>\r\n"
                    + "    <OperationType>T</OperationType>\r\n"
                    + "  </Header>\r\n"
                    + "  <Body>");
            headerWritten = true;
        }

        String currency = sheetName.substring(sheetName.length() - 3).toUpperCase();
        String appendixNumber = sheetName.replaceAll("[^0-9]", "");

        if (!openedCurrencies.contains(currency)) {
            writeXml("    <Appendix" + currency + ">");
            openedCurrencies.add(currency);
        }

        writeXml("      <Appendix" + appendixNumber + ">");

        int rowCount = sheet.getLastRowNum();
        if (rowCount == 0)
            return;

        List<String> allRateTags = new ArrayList<>();
        List<String> allBalanceTags = new ArrayList<>();
        List<String> dobandaTags = new ArrayList<>();

        for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {

            Row row = sheet.getRow(rowIndex);
            if (row == null)
                continue;
            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

            Cell cell_check = row.getCell(stringToInt.get(startColumn));
            if (isIgnoredColor(cell_check))
                continue;

            for (int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                String rawColCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow);
                if (rawColCode == null)
                    continue;

                colCode = rawColCode.trim();
                String cellValue = getCellValue(cell);

                if (cellValue == null || cellValue.equalsIgnoreCase("X") || cellValue.trim().isEmpty())
                    continue;

                String cellCode = appendixNumber + String.format("%03d", rowIndex) + colCode;
                String tagContent = "          <CellCode>" + cellCode + "</CellCode>\r\n"
                        + "          <CellValue>" + cellValue + "</CellValue>";

                switch (colCode) {
                    case "1":
                    case "3":
                    case "5":
                        allRateTags.add("        <Rate>\r\n" + tagContent + "\r\n        </Rate>");
                        break;

                    case "2":
                    case "4":
                    case "6":
                        allBalanceTags.add("        <Balance>\r\n" + tagContent + "\r\n        </Balance>");
                        break;
                }

                switch (rowCode) {
                    case "Dobanda anuala efectiva":
                        dobandaTags.add("        <AnnualAverageInterest>\r\n" + tagContent + "\r\n        </AnnualAverageInterest>");
                        break;
                }
            }
        }

        for (String rate : allRateTags) {
            writeXml(rate);
        }

        for (String balance : allBalanceTags) {
            writeXml(balance);
        }

        for (String dobanda : dobandaTags) {
            writeXml(dobanda);
        }

        writeXml("      </Appendix" + appendixNumber + ">");

        if (appendixNumber.contains("7")) {
            writeXml("    </Appendix" + currency + ">");
        }
    }





    public static void writeXmlSTASTV()throws SQLException
    {
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String sheetName = sheet.getSheetName();

        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);

        if (currentSheetName.equals("Formular1")){
            messageType = "STA-STV";
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns=\"http://www.bnr.ro/RAPDIR/schemas/message\">");
            writeXml("    <Header>");
            writeXml("        <SenderID>" + valueOfB271 + "</SenderID>");
            writeXml("        <ReferenceDate>" + refDate + "</ReferenceDate>");
            writeXml("        <SenderMsgID>"+ SenderMSGID +"</SenderMsgID>");
            writeXml("        <MessageType>STA-STV</MessageType>");
            writeXml("    </Header>");
            writeXml("    <Body>");
            writeXml("    	<Form code=\"anexa1\" instanceNo=\"1\">");
            writeXml("    		<Section code=\"fiCodeSection\">");
            writeXml("    			<Row code=\"1\">");
            writeXml("    				<Col code=\"fiCode\"/>");
            writeXml("    			</Row>");
            writeXml("    		</Section>");
        }

        if (currentSheetName.equals("Formular1")) {
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 4)
                return;
            writeXml("			<Section code=\"anexa1\">");

            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;

                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;


                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals("") || cellValue.trim().equals("."))
                        continue;

                    if(colCode.trim().equals("1"))
                        writeXml("    			<Row code=\""+ cellValue +"\">");
                    else if(colCode.trim().equals("2"))
                        writeXml("    				<Col code=\"codIsin\">"+ cellValue +"</Col>");
                    else if(colCode.trim().equals("3"))
                        writeXml("    				<Col code=\"codTipDetineri\">"+ cellValue + "</Col>");
                    else if(colCode.trim().equals("4"))
                        writeXml("    				<Col code=\"codTara\">"+ cellValue +"</Col>");
                    else if(colCode.trim().equals("5"))
                        writeXml("    				<Col code=\"codSectInstit\">"+ cellValue +"</Col>");
                    else if(colCode.trim().equals("6"))
                        writeXml("    				<Col code=\"codTipInvestitor\">"+ cellValue +"</Col>");
                    else if(colCode.trim().equals("7"))
                        writeXml("    				<Col code=\"codTipTranz\">"+ cellValue +"</Col>");
                    else if(colCode.trim().equals("8")) {
                        writeXml("    				<Col code=\"nrInstrumente\">"+ cellValue +"</Col>");
                        writeXml("    			</Row>");}

                }

            }
            writeXml("    		</Section>");
            writeXml("    	</Form>");

        }	if (currentSheetName.equals("Formular2")) {
        int rowCount = sheet.getLastRowNum();
        //System.out.println(rowCount);
        if(rowCount == 4)
            return;
        writeXml("    	<Form code=\"anexa2\" instanceNo=\"1\">");
        writeXml("    		<Section code=\"fiCodeSection\">");
        writeXml("    			<Row code=\"1\">");
        writeXml("    				<Col code=\"fiCode\">"+ valueOfB271 +"</Col>");
        writeXml("    			</Row>");
        writeXml("    		</Section>");
        writeXml("			<Section code=\"anexa2\">");

        for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
        {

            Row row = sheet.getRow(rowIndex);
            if(row == null)
                return;

            Cell cell_check = row.getCell(stringToInt.get(startColumn));
            if(isIgnoredColor(cell_check))
                continue;


            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);



                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                String cellValue = getCellValue(cell);


                if (cellValue.trim().equals("") || cellValue.trim().equals("."))
                    continue;

                if(colCode.trim().equals("1"))
                    writeXml("    			<Row code=\""+ cellValue +"\">");
                else if(colCode.trim().equals("2"))
                    writeXml("    				<Col code=\"codIsin\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("3"))
                    writeXml("    				<Col code=\"codTipDetineri\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("4"))
                    writeXml("    				<Col code=\"dataIncasare\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("5"))
                    writeXml("    				<Col code=\"valoareDividend\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("6")) {
                    writeXml("    				<Col code=\"monedaDividend\">"+ cellValue +"</Col>");
                    writeXml("    			</Row>");}

            }

        }	writeXml("    		</Section>");
        writeXml("    	</Form>");

    }	if (currentSheetName.equals("Formular4")) {
        int rowCount = sheet.getLastRowNum();
        //System.out.println(rowCount);
        if(rowCount == 8)
            return;
        writeXml("    	<Form code=\"anexa4\" instanceNo=\"1\">");
        writeXml("    		<Section code=\"fiCodeSection\">");
        writeXml("    			<Row code=\"1\">");
        writeXml("    				<Col code=\"fiCode\">"+ valueOfB271 +"</Col>");
        writeXml("    			</Row>");
        writeXml("    		</Section>");
        writeXml("			<Section code=\"anexa4\">");

        for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
        {

            Row row = sheet.getRow(rowIndex);
            if(row == null)
                return;

            Cell cell_check = row.getCell(stringToInt.get(startColumn));
            if(isIgnoredColor(cell_check))
                continue;


            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);



                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                String cellValue = getCellValue(cell);


                if (cellValue.trim().equals("") || cellValue.trim().equals("."))
                    continue;

                if(colCode.trim().equals("1"))
                    writeXml("    			<Row code=\""+ cellValue +"\">");
                else if(colCode.trim().equals("2"))
                    writeXml("    				<Col code=\"codIsin\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("3"))
                    writeXml("    				<Col code=\"dataEmisiune\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("4"))
                    writeXml("    				<Col code=\"valoareEmisiune\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("5"))
                    writeXml("    				<Col code=\"monedaEmisiune\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("6"))
                    writeXml("    				<Col code=\"tipRascumparare\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("7"))
                    writeXml("    				<Col code=\"frecventaPlata\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("8"))
                    writeXml("    				<Col code=\"monedaRascumparare\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("9"))
                    writeXml("    				<Col code=\"dataPrima\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("10"))
                    writeXml("    				<Col code=\"dataUltima\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("11"))
                    writeXml("    				<Col code=\"tipCupon\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("12"))
                    writeXml("    				<Col code=\"dataPrimaCupon\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("13"))
                    writeXml("    				<Col code=\"dataUltimaCupon\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("14"))
                    writeXml("    				<Col code=\"frecventaCupon\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("15"))
                    writeXml("    				<Col code=\"monedaCupon\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("16"))
                    writeXml("    				<Col code=\"rataCupon\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("17"))
                    writeXml("    				<Col code=\"soldEmisiune\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("18"))
                    writeXml("    				<Col code=\"pretEmisiune\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("19"))
                    writeXml("    				<Col code=\"valoareNominala\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("20"))
                    writeXml("    				<Col code=\"primaCall\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("21")) {
                    writeXml("    				<Col code=\"piataCapital\">"+ cellValue +"</Col>");
                    writeXml("    			</Row>");}

            }

        }	writeXml("    		</Section>");
        writeXml("    	</Form>");

    }if (currentSheetName.equals("Formular5")) {
        int rowCount = sheet.getLastRowNum();
        //System.out.println(rowCount);
        if(rowCount == 9)
            return;
        writeXml("    	<Form code=\"anexa5\" instanceNo=\"1\">");
        writeXml("    		<Section code=\"fiCodeSection\">");
        writeXml("    			<Row code=\"1\">");
        writeXml("    				<Col code=\"fiCode\">"+ valueOfB271 +"</Col>");
        writeXml("    			</Row>");
        writeXml("    		</Section>");
        writeXml("			<Section code=\"anexa5\">");

        for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
        {

            Row row = sheet.getRow(rowIndex);
            if(row == null)
                return;


            Cell cell_check = row.getCell(stringToInt.get(startColumn));
            if(isIgnoredColor(cell_check))
                continue;


            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);



                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                String cellValue = getCellValue(cell);


                if (cellValue.trim().equals("") || cellValue.trim().equals("."))
                    continue;

                if(colCode.trim().equals("1"))
                    writeXml("    			<Row code=\""+ cellValue +"\">");
                else if(colCode.trim().equals("2"))
                    writeXml("    				<Col code=\"codRef\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("3"))
                    writeXml("    				<Col code=\"codTip\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("4"))
                    writeXml("    				<Col code=\"codTaraOrg\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("5"))
                    writeXml("    				<Col code=\"codSectorInst\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("6"))
                    writeXml("    				<Col code=\"codTipInvest\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("7"))
                    writeXml("    				<Col code=\"soldDetineriNom\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("8"))
                    writeXml("    				<Col code=\"soldDetineriPiata\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("9"))
                    writeXml("    				<Col code=\"dividendIncasat\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("10"))
                    writeXml("    				<Col code=\"dataIncasareDividend\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("11"))
                    writeXml("    				<Col code=\"monedaDividend\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("12"))
                    writeXml("    				<Col code=\"procentInstrumente\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("13"))
                    writeXml("    				<Col code=\"dataEmisiune\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("14")) {
                    writeXml("    				<Col code=\"dataRascumparare\">"+ cellValue +"</Col>");
                    writeXml("    			</Row>");}
            }

        }	writeXml("    		</Section>");
        writeXml("    	</Form>");

    }if (currentSheetName.equals("Persoane")) {
        int rowCount = sheet.getLastRowNum();
        //System.out.println(rowCount);
        if(rowCount == 4)
            return;
        writeXml("    	<Form code=\"genFormPersContact\" instanceNo=\"1\">");

        for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
        {

            Row row = sheet.getRow(rowIndex);
            if(row == null)
                return;
            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

            if (rowCode == null || rowCode.trim().equals(""))
                continue;
            if(rowCode.trim().equals("Avizat")) {
                writeXml("			<Section code=\"genSectionAvizat\">");
            }else if(rowCode.trim().equals("Intocmit")) {
                writeXml("			<Section code=\"genSectionIntocmit\">");
            }

            Cell cell_check = row.getCell(stringToInt.get(startColumn));
            if(isIgnoredColor(cell_check))
                continue;



            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);



                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                String cellValue = getCellValue(cell);


                if (cellValue.trim().equals("") || cellValue.trim().equals("."))
                    continue;

                if(colCode.trim().equals("1"))
                    writeXml("    			<Row code=\""+ cellValue +"\">");
                else if(colCode.trim().equals("2"))
                    writeXml("    				<Col code=\"numePrenume\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("3"))
                    writeXml("    				<Col code=\"functie\">"+ cellValue + "</Col>");
                else if(colCode.trim().equals("4"))
                    writeXml("    				<Col code=\"telefon\">"+ cellValue +"</Col>");
                else if(colCode.trim().equals("5")) {
                    writeXml("    				<Col code=\"email\">"+ cellValue +"</Col>");
                    writeXml("    			</Row>");
                    writeXml("    		</Section>");}
            }

        }
        writeXml("    	</Form>");

    }


    }
    public static void writeXmlPVL110()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String xsd;
        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);



        if (currentSheetName.equals("Banks") && valueOfB271.equals("L")) {
            messageType = "PVL110L-01";
            xsd = "PVL110L-01.xsd";
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xsi:schemaLocation=\"http://extranet.bnr.ro " + xsd + "\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            writeXml("    <Header>");
            writeXml("        <SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("        <SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("        <MessageType>" + messageType + "</MessageType>");
            writeXml("        <RefDate>" + refDate + "</RefDate>");
            writeXml("        <SenderMessageId>000</SenderMessageId>");
            writeXml("        <OperationType>T</OperationType>");
            writeXml("    </Header>");
            writeXml("    <Body>");
            writeXml("        <IdMessageType>"+ valueOfB271 +"</IdMessageType>");
        } else if (currentSheetName.equals("Banks") && valueOfB271.equals("V")){
            messageType = "PVL110V-01";
            xsd = "PVL110V-01.xsd";
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xsi:schemaLocation=\"http://extranet.bnr.ro " + xsd + "\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            writeXml("    <Header>");
            writeXml("        <SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("        <SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("        <MessageType>" + messageType + "</MessageType>");
            writeXml("        <RefDate>" + refDate + "</RefDate>");
            writeXml("        <SenderMessageId>000</SenderMessageId>");
            writeXml("        <OperationType>T</OperationType>");
            writeXml("    </Header>");
            writeXml("    <Body>");
            writeXml("        <IdMessageType>"+ valueOfB271 +"</IdMessageType>");
        }


        if (currentSheetName.equals("Banks")) {
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 4)
                return;

            writeXml("        <Banks>");
            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    continue;

                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;

                writeXml("			<Transaction>");

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals(""))
                        continue;
                    if(colCode.trim().equals("20"))
                        writeXml("					<TransactionId>" + cellValue + "</TransactionId>");
                    else if (colCode.trim().equals("30"))
                        writeXml("					<OperationType>" + cellValue + "</OperationType>");
                    else if (colCode.trim().equals("40"))
                        writeXml("					<Cntpart>" + cellValue + "</Cntpart>");
                    else if (colCode.trim().equals("50"))
                        writeXml("					<Currency1>" + cellValue + "</Currency1>");
                    else if (colCode.trim().equals("60"))
                        writeXml("					<Currency2>" + cellValue + "</Currency2>");
                    else if (colCode.trim().equals("70"))
                        writeXml("					<ExchangeRate>" + cellValue + "</ExchangeRate>");
                    else if (colCode.trim().equals("80"))
                        writeXml("					<Amount1>" + cellValue + "</Amount1>");
                    else if (colCode.trim().equals("90"))
                        writeXml("					<Amount2>" + cellValue + "</Amount2>");
                    else if (colCode.trim().equals("100"))
                        writeXml("					<TransactionType>" + cellValue + "</TransactionType>");
                    else if (colCode.trim().equals("110") && !cellValue.trim().equals(""))
                        writeXml("					<SwapReference>" + cellValue + "</SwapReference>");
                    else if (colCode.trim().equals("120"))
                        writeXml("					<ExchangeDate>" + cellValue + "</ExchangeDate>");
                }
                writeXml("				</Transaction>");

            }
            writeXml("		</Banks>");
        }
        if (currentSheetName.equals("Clients")) {
            int rowcnt = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowcnt == 4)
                return;
            writeXml("			<Clients>");

            for(int rowIndex = startRow; rowIndex <= rowcnt; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    continue;

                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;


                writeXml("				<Transaction>");

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals(""))
                        continue;
                    if(colCode.trim().equals("20"))
                        writeXml("					<TransactionId>" + cellValue + "</TransactionId>");
                    else if (colCode.trim().equals("30"))
                        writeXml("					<OperationType>" + cellValue + "</OperationType>");
                    else if (colCode.trim().equals("50"))
                        writeXml("					<ClientName>" + cellValue + "</ClientName>");
                    else if (colCode.trim().equals("60"))
                        writeXml("					<ClientType>" + cellValue + "</ClientType>");
                    else if (colCode.trim().equals("70"))
                        writeXml("					<Currency1>" + cellValue + "</Currency1>");
                    else if (colCode.trim().equals("80"))
                        writeXml("					<Currency2>" + cellValue + "</Currency2>");
                    else if (colCode.trim().equals("90"))
                        writeXml("					<ExchangeRate>" + cellValue + "</ExchangeRate>");
                    else if (colCode.trim().equals("100"))
                        writeXml("					<Amount1>" + cellValue + "</Amount1>");
                    else if (colCode.trim().equals("110"))
                        writeXml("					<Amount2>" + cellValue + "</Amount2>");
                    else if (colCode.trim().equals("120"))
                        writeXml("					<TransactionType>" + cellValue + "</TransactionType>");
                    else if (colCode.trim().equals("130") && cellValue.trim().equals(""))
                        writeXml("					<SwapReference>" + cellValue + "</SwapReference>");
                    else if (colCode.trim().equals("140"))
                        writeXml("					<ExchangeDate>" + cellValue + "</ExchangeDate>");
                }
                writeXml("				</Transaction>");
            }
            writeXml("			</Clients>");
        }

    }


    public static void writeXmlPMI120()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String xsd;
        int cnt = 0;

        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);



        messageType = "PMI120-01";
        xsd = "PMI120-01.xsd";
        writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xsi:schemaLocation=\"http://extranet.bnr.ro " + xsd + "\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        writeXml("    <Header>");
        writeXml("        <SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
        writeXml("        <SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
        writeXml("        <MessageType>" + messageType + "</MessageType>");
        writeXml("        <RefDate>" + refDate + "</RefDate>");
        writeXml("        <SenderMessageId>000</SenderMessageId>");
        writeXml("        <OperationType>T</OperationType>");
        writeXml("    </Header>");
        writeXml("    <Body>");


        if (currentSheetName.equals("Transactions")) {
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;

            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;


                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;
                int counter = 1;


                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals(""))
                        return;
                    if(cnt == 0) {
                        writeXml("			<Transaction>");
                        cnt++;
                    }

                    if (colCode.trim().equals("10"))
                        writeXml("					<TransactionId>" + cellValue + "</TransactionId>");
                    else if (colCode.trim().equals("20"))
                        writeXml("					<Corespondent>" + cellValue + "</Corespondent>");
                    else if (colCode.trim().equals("30"))
                        writeXml("					<BeginDate>" + cellValue + "</BeginDate>");
                    else if (colCode.trim().equals("40"))
                        writeXml("					<EndDate>" + cellValue + "</EndDate>");
                    else if (colCode.trim().equals("50"))
                        writeXml("					<Amount>" + cellValue + "</Amount>");
                    else if (colCode.trim().equals("60"))
                        writeXml("					<Interest>" + cellValue + "</Interest>");
                    else if (colCode.trim().equals("70"))
                        writeXml("					<DepositType>" + cellValue + "</DepositType>");
                    else if (colCode.trim().equals("80"))
                        writeXml("					<Maturity>" + cellValue + "</Maturity>");
                }
                writeXml("			</Transaction>");
                cnt = 0;
            }

        }
    }


   public static void writeXmlRPS100() throws SQLException {

       validationErrors.clear();

       String variabilaUnitateBancara = reportingEntity;
       String colCode;

       if (!(refDate.endsWith("06-30") || refDate.endsWith("12-31"))) {
           validationErrors.add("Header.RefDate nu este corecta: " + refDate);
       }

       if (currentSheetName.equals("Main")) {

           writeXml(
                   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                           "<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro RPS100-01.xsd\">\n" +
                           "    <Header>\n" +
                           "        <SenderId>" + reportingEntity + "</SenderId>\n" +
                           "        <SendingDate>2026-01-11</SendingDate>\n" +
                           "        <MessageType>RPS100-01</MessageType>\n" +
                           "        <RefDate>" + refDate + "</RefDate>\n" +
                           "        <SenderMessageId>" + SenderMSGID + "</SenderMessageId>\n" +
                           "        <OperationType>T</OperationType>\n" +
                           "    </Header>\n" +
                           "    <Body>"
           );

           int rowCount = sheet.getLastRowNum();

           for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {

               Row row = sheet.getRow(rowIndex);
               if (row == null) continue;

               Cell cell_check = row.getCell(stringToInt.get(startColumn));
               if (isIgnoredColor(cell_check)) continue;

               for (int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++) {

                   Cell cell = row.getCell(columnIndex);

                   colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                   String cellValue = getCellValue(cell);

                   if (cellValue.trim().isEmpty()) {
                       validationErrors.add("Main row " + rowIndex + " col " + colCode + " este gol");
                       continue;
                   }

                   if (colCode.equals("10")) {

                       if (!cellValue.equals("C")) {
                           validationErrors.add("Body.Type incorect la row " + rowIndex);
                       }

                       writeXml("        <Type>" + cellValue + "</Type>");
                   }
               }
           }
       }

       if (currentSheetName.equals("Form11")) {

           writeXml("        <Form11>");

           int rowCount = sheet.getLastRowNum();

           for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {

               String buyerNo = null;
               String buyerUID = null;
               String buyerName = null;
               String buyerLEI = null;

               Row row = sheet.getRow(rowIndex);
               if (row == null) continue;

               writeXml("            <Item>");

               Cell cell_check = row.getCell(stringToInt.get(startColumn));
               if (isIgnoredColor(cell_check)) continue;

               for (int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++) {

                   Cell cell = row.getCell(columnIndex);

                   colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                   String cellValue = getCellValue(cell);

                   if (cellValue.trim().isEmpty()) continue;

                   if (colCode.equals("10")) {

                       buyerNo = cellValue;

                       if (!buyerNo.matches("\\d{1,6}")) {
                           validationErrors.add("Form11 row " + rowIndex + " BuyerNo invalid: " + buyerNo);
                       }

                       if (!buyerNos.add(buyerNo)) {
                           validationErrors.add("Form11 BuyerNo duplicat: " + buyerNo);
                       }

                       writeXml("                <BuyerNo>" + cellValue + "</BuyerNo>");
                   }

                   if (colCode.equals("20")) {

                       buyerUID = cellValue;

                       if (cellValue.length() > 50) {
                           validationErrors.add("Form11 row " + rowIndex + " BuyerUID prea lung");
                       }

                       if (!buyerUIDs.add(buyerUID)) {
                           validationErrors.add("Form11 BuyerUID duplicat: " + buyerUID);
                       }

                       form11BuyerUIDs.add(buyerUID);

                       writeXml("                <BuyerUID>" + cellValue + "</BuyerUID>");
                   }

                   if (colCode.equals("30")) {
                       buyerName = cellValue;
                       writeXml("                <BuyerName>" + cellValue + "</BuyerName>");
                   }

                   if (colCode.equals("40")) {
                       buyerLEI = cellValue;
                       writeXml("                <BuyerLEI>" + cellValue + "</BuyerLEI>");
                   }

                   if (colCode.equals("50"))
                       writeXml("                <BuyerAddress>" + cellValue + "</BuyerAddress>");

                   if (colCode.equals("60"))
                       writeXml("                <RepLEI>" + cellValue + "</RepLEI>");

                   if (colCode.equals("70"))
                       writeXml("                <RepAddress>" + cellValue + "</RepAddress>");
               }

               if (buyerName == null && buyerLEI == null) {
                   validationErrors.add("Form11 row " + rowIndex + " trebuie BuyerName sau BuyerLEI");
               }

               writeXml("            </Item>");
           }

           writeXml("        </Form11>");
       }

       if (currentSheetName.equals("Form12")) {

           writeXml("        <Form12>");

           int rowCount = sheet.getLastRowNum();

           for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {

               String buyerUID = null;
               String entType = null;
               String entUID = null;
               String entName = null;

               Row row = sheet.getRow(rowIndex);
               if (row == null) continue;

               writeXml("            <Item>");

               Cell cell_check = row.getCell(stringToInt.get(startColumn));
               if (isIgnoredColor(cell_check)) continue;

               for (int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++) {

                   Cell cell = row.getCell(columnIndex);

                   colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                   String cellValue = getCellValue(cell);

                   if (cellValue.trim().isEmpty()) continue;

                   if (colCode.equals("10")) {

                       buyerUID = cellValue;

                       if (!form11BuyerUIDs.contains(buyerUID)) {
                           validationErrors.add("Form12 row " + rowIndex + " BuyerUID nu exista in Form11: " + buyerUID);
                       }

                       writeXml("                <BuyerUID>" + cellValue + "</BuyerUID>");
                   }

                   if (colCode.equals("20")) {

                       entType = cellValue;

                       if (!(entType.equals("D") || entType.equals("M"))) {
                           validationErrors.add("Form12 row " + rowIndex + " EntType invalid: " + entType);
                       }

                       writeXml("                <EntType>" + cellValue + "</EntType>");
                   }

                   if (colCode.equals("30")) {
                       entName = cellValue;
                       writeXml("                <EntName>" + cellValue + "</EntName>");
                   }

                   if (colCode.equals("40")) {
                       entUID = cellValue;
                       writeXml("                <EntUID>" + cellValue + "</EntUID>");
                   }
               }

               if (entName == null && entUID == null) {
                   validationErrors.add("Form12 row " + rowIndex + " trebuie EntName sau EntUID");
               }

               String key = buyerUID + "|" + entType + "|" + entUID;
               if (!form12Unique.add(key)) {
                   validationErrors.add("Form12 combinatie duplicata: " + key);
               }

               writeXml("            </Item>");
           }

           writeXml("        </Form12>");
       }

       if (currentSheetName.equals("Form2")) {

           writeXml("        <Form2>");

           int rowCount = sheet.getLastRowNum();

           for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {

               String buyerUID = null;
               String transNo = null;
               BigDecimal totalBal = BigDecimal.ZERO;
               BigDecimal totalQty = BigDecimal.ZERO;
               BigDecimal avgBal = BigDecimal.ZERO;

               Row row = sheet.getRow(rowIndex);
               if (row == null) continue;

               writeXml("            <Item>");

               Cell cell_check = row.getCell(stringToInt.get(startColumn));
               if (isIgnoredColor(cell_check)) continue;

               for (int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++) {

                   Cell cell = row.getCell(columnIndex);

                   colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                   String cellValue = getCellValue(cell);

                   if (cellValue.trim().isEmpty()) continue;

                   if (colCode.equals("10")) {
                       transNo = cellValue;
                       writeXml("                <TransNo>" + cellValue + "</TransNo>");
                   }

                   if (colCode.equals("20")) {

                       buyerUID = cellValue;

                       if (!form11BuyerUIDs.contains(buyerUID)) {
                           validationErrors.add("Form2 row " + rowIndex + " BuyerUID nu exista in Form11");
                       }

                       writeXml("                <BuyerUID>" + cellValue + "</BuyerUID>");
                   }

                   if (colCode.equals("40")) {
                       totalBal = new BigDecimal(cellValue);
                       writeXml("                <TotalNPLBal>" + cellValue + "</TotalNPLBal>");
                   }

                   if (colCode.equals("50")) {
                       totalQty = new BigDecimal(cellValue);
                       writeXml("                <TotalNplQty>" + cellValue + "</TotalNplQty>");
                   }

                   if (colCode.equals("60")) {
                       avgBal = new BigDecimal(cellValue);
                       writeXml("                <AvgNPLBal>" + cellValue + "</AvgNPLBal>");
                   }

                   if (colCode.equals("70"))
                       writeXml("                <InclConsumer>" + cellValue + "</InclConsumer>");

                   if (colCode.equals("80")) {

                       String[] parts = cellValue.split(",");
                       Set<String> allowed = new HashSet<>(Arrays.asList("A", "B", "C", "D", "E", "F", "G", "H"));
                       Set<String> seen = new HashSet<>();

                       for (String p : parts) {
                           String trimmed = p.trim();  // ELIMINA SPATIILE

                           if (!allowed.contains(trimmed)) {
                               validationErrors.add("Form2 row " + rowIndex + " CollateralTypes invalid: " + cellValue);
                           }

                           if (!seen.add(trimmed)) {
                               validationErrors.add("Form2 row " + rowIndex + " CollateralTypes duplicat: " + cellValue);
                           }
                       }

                       writeXml("                <CollateralTypes>" + cellValue.replace(" ", "") + "</CollateralTypes>");
                   }
               }

               if (totalQty.compareTo(BigDecimal.ZERO) != 0) {
                   BigDecimal calculated = totalBal.divide(totalQty, 2, RoundingMode.HALF_UP);
                   if (calculated.compareTo(avgBal) != 0) {
                       validationErrors.add("Form2 row " + rowIndex + " AvgNPLBal incorect: " +
                               totalBal + " / " + totalQty + " = " + calculated +
                               " (expected: " + avgBal + ")");
                   }

                   writeXml("            </Item>");
               }

               writeXml("        </Form2>");
           }

           if (!validationErrors.isEmpty()) {
               hasErrors = true;
               writeXml("        <ValidationErrors>");

               for (String err : validationErrors) {
                   writeXml("            <Error>" + err + "</Error>");
               }

               writeXml("        </ValidationErrors>");
           }

       }
   }
    public static void writeXmlARM()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String xsd;
        int cnt = 0;

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        String executionDateTime = LocalDateTime.now().format(formatter);
        Instant now = Instant.now();
        String currentDate = now.toString();
        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);

        int SenderMSGIDInt = Integer.parseInt(SenderMSGID);

        writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        writeXml("<BizData xmlns=\"urn:iso:std:iso:20022:tech:xsd:head.003.001.01\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"urn:iso:std:iso:20022:tech:xsd:head.003.001.01 head.003.001.01.xsd\">");
        writeXml("<Hdr>");
        writeXml("		<AppHdr xmlns=\"urn:iso:std:iso:20022:tech:xsd:head.001.001.01\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"urn:iso:std:iso:20022:tech:xsd:head.001.001.01 head.001.001.01_ESMAUG_1.0.0.xsd\">");
        writeXml("<Fr>\r\n"
                + "				<OrgId>\r\n"
                + "					<Id>\r\n"
                + "						<OrgId>\r\n"
                + "							<Othr>\r\n"
                + "								<Id>"+ LEICode +"</Id>\r\n"
                + "								<SchmeNm>\r\n"
                + "									<Prtry>LEI</Prtry>\r\n"
                + "								</SchmeNm>\r\n"
                + "							</Othr>\r\n"
                + "						</OrgId>\r\n"
                + "					</Id>\r\n"
                + "				</OrgId>\r\n"
                + "			</Fr>\r\n"
                + "			<To>\r\n"
                + "				<OrgId>\r\n"
                + "					<Id>\r\n"
                + "						<OrgId>\r\n"
                + "							<Othr>\r\n"
                + "								<Id>RO</Id>\r\n"
                + "							</Othr>\r\n"
                + "						</OrgId>\r\n"
                + "					</Id>\r\n"
                + "				</OrgId>\r\n"
                + "			</To>\r\n"
                + "			<BizMsgIdr>BID_DATTRA_BNRO_"+ SenderMSGID +"_0_"+ ( SenderMSGIDInt - 1) +""+"</BizMsgIdr>\r\n"
                + "			<MsgDefIdr>auth.016.001.01</MsgDefIdr>\r\n"
                + "			<CreDt>"+ currentDate +"</CreDt>\r\n"
                + "		</AppHdr>\r\n"
                + "</Hdr>");


        if (currentSheetName.equals("Transactions")) {
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;
            writeXml("	<Pyld>");
            writeXml("		<Document xmlns=\"urn:iso:std:iso:20022:tech:xsd:auth.016.001.01\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"urn:iso:std:iso:20022:tech:xsd:auth.016.001.01 auth.016.001.01_ESMAUG_Reporting_1.1.0.xsd\">");
            writeXml("			<FinInstrmRptgTxRpt>");

            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;
                String armCCY = getCellValue(row.getCell(31));
                String pricCCY = getCellValue(row.getCell(34));
                String ctrDec = getCellValue(row.getCell(59));
                String ctrExec = getCellValue(row.getCell(61));
                String ExCd = getCellValue(row.getCell(65));
                String SctFin = getCellValue(row.getCell(66));
                String tradPclMtchId = getCellValue(row.getCell(3));
                String exctgID = getCellValue(row.getCell(58));
                String dateCreated = getCellValue(row.getCell(0));
                String reportStatus = getCellValue(row.getCell(1));
                String transactionRefNr = getCellValue(row.getCell(2));
                String tradingVenueTrnIdCode = getCellValue(row.getCell(3));
                String executingEntityIdCode = getCellValue(row.getCell(4));
                String investmentFirm = getCellValue(row.getCell(5));
                String submittingEntityIdCode = getCellValue(row.getCell(6));
                String buyerIdCode = getCellValue(row.getCell(7));
                String countryBranchBuyer = getCellValue(row.getCell(8));
                String buyerFirstName = getCellValue(row.getCell(9));
                String buyerSurname = getCellValue(row.getCell(10));
                String buyerDateOfBirth = getCellValue(row.getCell(11));
                String buyerDmCode = getCellValue(row.getCell(12));
                String buyDmFirstName = getCellValue(row.getCell(13));
                String buyDmSurname = getCellValue(row.getCell(14));
                String buyDmDateOfBirth = getCellValue(row.getCell(15));
                String sellerIdCode = getCellValue(row.getCell(16));
                String countryBranchSeller = getCellValue(row.getCell(17));
                String sellerFirstName = getCellValue(row.getCell(18));
                String sellerSurname = getCellValue(row.getCell(19));
                String sellerDateOfBirth = getCellValue(row.getCell(20));
                String sellerDmCode = getCellValue(row.getCell(21));
                String sellerDmFirstName = getCellValue(row.getCell(22));
                String sellerDmSurname = getCellValue(row.getCell(23));
                String sellerDmDateOfBirth = getCellValue(row.getCell(24));
                String transmisOrderIndicator = getCellValue(row.getCell(25));
                String transmitFirmIdCodeBuyer = getCellValue(row.getCell(26));
                String transmitFirmIdCodeSeller = getCellValue(row.getCell(27));
                String tradingDateTime = getCellValue(row.getCell(28));
                String tradingCapacity = getCellValue(row.getCell(29));
                String quantity = getCellValue(row.getCell(30));
                String quantityCurrency = getCellValue(row.getCell(31));
                String derivativeNotionalIncrDecr = getCellValue(row.getCell(32));
                String price = getCellValue(row.getCell(33));
                String priceCurrency = getCellValue(row.getCell(34));
                String netAmount = getCellValue(row.getCell(35));
                String venue = getCellValue(row.getCell(36));
                String countryBranchMembership = getCellValue(row.getCell(37));
                String tradPlcMtchgldId = getCellValue(row.getCell(38));
                String upFrontPayment = getCellValue(row.getCell(39));
                String upFrontPaymentCurrency = getCellValue(row.getCell(40));
                String complexTradeComponentId = getCellValue(row.getCell(41));
                String instrumentIdCode = getCellValue(row.getCell(42));
                String instrumentFullName = getCellValue(row.getCell(43));
                String instrumentClassification = getCellValue(row.getCell(44));
                String notionalCurrency1 = getCellValue(row.getCell(45));
                String notionalCurrency2 = getCellValue(row.getCell(46));
                String priceMultiplier = getCellValue(row.getCell(47));
                String underlyingInstrumentCode = getCellValue(row.getCell(48));
                String underlyingIndexName = getCellValue(row.getCell(49));
                String termUnderlyingIndex = getCellValue(row.getCell(50));
                String optionType = getCellValue(row.getCell(51));
                String strikePrice = getCellValue(row.getCell(52));
                String strikePriceCurrency = getCellValue(row.getCell(53));
                String optionExerciseStyle = getCellValue(row.getCell(54));
                String maturityDate = getCellValue(row.getCell(55));
                String expiryDateFinInstr = getCellValue(row.getCell(56));
                String deliveryType = getCellValue(row.getCell(57));
                String investDecisionWithinFirm = getCellValue(row.getCell(58));
                String countryBranchPersDecision = getCellValue(row.getCell(59));
                String executionWithinFirm = getCellValue(row.getCell(60));
                String countryBranchPersExecution = getCellValue(row.getCell(61));
                String waiverIndicator = getCellValue(row.getCell(62));
                String shortSellingIndicator = getCellValue(row.getCell(63));
                String otcPostTradeIndicator = getCellValue(row.getCell(64));
                String commodityDerivIndicator = getCellValue(row.getCell(65));
                String securitiesFinTrnIndicator = getCellValue(row.getCell(66));
                String codeTypeBuyer = getCellValue(row.getCell(67));
                String codeTypeSeller = getCellValue(row.getCell(68));
                String codeTypeBuyerDecision = getCellValue(row.getCell(69));
                String codeTypeSellerDecision = getCellValue(row.getCell(70));
                String quantityType = getCellValue(row.getCell(71));
                String priceUnitType = getCellValue(row.getCell(72));
                String RskRdcgTx = getCellValue(row.getCell(73));

                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;
                int counter = 1;


                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals("") && colCode.equals("10"))
                        return;


                    if (colCode.trim().equals("30")) {
                        writeXml("        <Tx>");
                        writeXml("            <New>");
                        writeXml("                <TxId>" + cellValue + "</TxId>");
                    }
                    else if (colCode.trim().equals("50")) {
                        writeXml("                <ExctgPty>" + cellValue + "</ExctgPty>");
                    }
                    else if (colCode.trim().equals("60")) {
                        writeXml("                <InvstmtPtyInd>" + cellValue + "</InvstmtPtyInd>");
                    }
                    else if (colCode.trim().equals("70")) {
                        writeXml("                <SubmitgPty>" + cellValue + "</SubmitgPty>");
                    }
                    else if (colCode.trim().equals("80")) {
                        writeXml("                <Buyr>");
                        writeXml("                    <AcctOwnr>");
                        writeXml("                        <Id>");
                        writeXml("                            <LEI>" + cellValue + "</LEI>");
                        writeXml("                        </Id>");
                        writeXml("                <CtryOfBrnch>"+ countryBranchBuyer +"</CtryOfBrnch>");
                        writeXml("                    </AcctOwnr>");
                        writeXml("                </Buyr>");
                    }
                    else if (colCode.trim().equals("170")) {
                        writeXml("                <Sellr>");
                        writeXml("                    <AcctOwnr>");
                        writeXml("                        <Id>");
                        writeXml("                            <LEI>" + cellValue + "</LEI>");
                        writeXml("                        </Id>");
                        writeXml("                <CtryOfBrnch>"+ countryBranchSeller +"</CtryOfBrnch>");
                        writeXml("                    </AcctOwnr>");
                        writeXml("                </Sellr>");
                    }
                    else if (colCode.trim().equals("260")) {
                        writeXml("                <OrdrTrnsmssn>");
                        writeXml("                    <TrnsmssnInd>" + cellValue + "</TrnsmssnInd>");
                        writeXml("                </OrdrTrnsmssn>");
                    }
                    else if (colCode.trim().equals("290")) {
                        writeXml("                <Tx>");
                        if (cell != null) {
                            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                                // Excel date -> Java date
                                Date excelDate = cell.getDateCellValue();
                                Instant instant = excelDate.toInstant();

                                // Formatare �n stil ISO 8601 cu sufixul Z pentru UTC
                                DateTimeFormatter formattter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss'Z'").withZone(ZoneOffset.UTC);
                                cellValue = formattter.format(instant); // Ex: 2025-07-24T08:54:00Z
                            } else {
                                // fallback pentru alte tipuri
                                cellValue = cell.toString().trim();
                            }
                            writeXml("                    <TradDt>" + cellValue + "</TradDt>");
                        }

                    }
                    else if (colCode.trim().equals("300")) {
                        writeXml("                    <TradgCpcty>" + cellValue + "</TradgCpcty>");
                    }
				/*else if (colCode.trim().equals("310")) {
					writeXml("                    <Qty>");
					if(!armCCY.trim().equals("") && !priceUnitType.equals("MONE"))
						writeXml("                        <NmnlVal Ccy=\"" + armCCY + "\">" + cellValue + "</NmnlVal>");
					else if(!armCCY.trim().equals("") && priceUnitType.equals("MONE")) writeXml("                        <MntryVal Ccy=\"" + armCCY + "\">" + quantity + "</MntryVal>");

					else if (!priceUnitType.equals("MONE") && armCCY.trim().equals(""))
						writeXml("                        <NmnlVal>" + cellValue + "</NmnlVal>");
					else writeXml("                        <NmnlVal>" + cellValue + "</NmnlVal>");
					writeXml("                    </Qty>");
				}*/
                    else if (colCode.trim().equals("310")) {
                        writeXml("                    <Qty>");

                        boolean hasCcy = !armCCY.trim().isEmpty();
                        boolean isMoney = priceUnitType.equals("MONE");

                        String tagName = isMoney ? "MntryVal" : "NmnlVal";
                        String value = isMoney ? quantity : cellValue;

                        if (hasCcy) {
                            writeXml("                        <" + tagName + " Ccy=\"" + armCCY + "\">" + value + "</" + tagName + ">");
                        } else {
                            writeXml("                        <" + tagName + ">" + value + "</" + tagName + ">");
                        }

                        writeXml("                    </Qty>");
                    }
                    else if (colCode.trim().equals("340")) {
                        writeXml("                    <Pric>");
                        writeXml("                        <Pric>");
                        if (quantityType.equals("MONE") && priceUnitType.equals("MONE"))
                            writeXml("							<MntryVal>\r\n"
                                    + "								<Amt Ccy=\""+ priceCurrency +"\">"+ cellValue +"</Amt>\r\n"
                                    + "							</MntryVal>");
                        else writeXml("                                <Pctg>" + cellValue + "</Pctg>");
                        writeXml("                        </Pric>");
                        writeXml("                    </Pric>");
                    }
                    else if (colCode.trim().equals("360") && !cellValue.trim().equals("")) {
                        writeXml("                    <NetAmt>" + cellValue + "</NetAmt>");
                    }
                    else if (colCode.trim().equals("370")) {
                        writeXml("                    <TradVn>" + cellValue + "</TradVn>");
                        if(!upFrontPayment.equals("")) {
                            writeXml("							<UpFrntPmt>\r\n"
                                    + "								<Amt Ccy=\""+ upFrontPaymentCurrency +"\">"+ upFrontPayment +"</Amt>\r\n"
                                    + "								<Sgn>"+ complexTradeComponentId +"</Sgn>\r\n"
                                    + "							</UpFrntPmt>");}

                    }
                    else if (colCode.trim().equals("380") && !cellValue.trim().equals("")) {
                        writeXml("                    <CtryOfBrnch>" + cellValue + "</CtryOfBrnch>");
                        if(!tradPlcMtchgldId.trim().equals(""))
                            writeXml("                <TradPlcMtchgId>"+tradPlcMtchgldId+"</TradPlcMtchgId>");
                    }
                    else if (colCode.trim().equals("390")) {
                        writeXml("                </Tx>");

                    }
                    else if (colCode.trim().equals("430"))
                        writeXml("						<FinInstrm>\r\n"
                                + "							<Id>"+ cellValue +"</Id>\r\n"
                                + "						</FinInstrm>");
                    else if (colCode.trim().equals("590")) {
                        writeXml("						<InvstmtDcsnPrsn>\r\n"
                                + "							<Prsn>\r\n"
                                + "								<CtryOfBrnch>"+ ctrDec +"</CtryOfBrnch>\r\n"
                                + "								<Othr>\r\n"
                                + "									<Id>"+ cellValue +"</Id>");
                    }
                    else if (colCode.trim().equals("600"))
                        writeXml("									<SchmeNm>\r\n"
                                + "										<Cd>"+ commodityDerivIndicator +"</Cd>\r\n"
                                + "									</SchmeNm>\r\n"
                                + "								</Othr>\r\n"
                                + "							</Prsn>\r\n"
                                + "						</InvstmtDcsnPrsn>");
                    else if (colCode.trim().equals("660"))
                        writeXml("						<ExctgPrsn>\r\n"
                                + "							<Prsn>\r\n"
                                + "								<CtryOfBrnch>"+ ctrExec +"</CtryOfBrnch>\r\n"
                                + "								<Othr>\r\n"
                                + "									<Id>"+ exctgID +"</Id>\r\n"
                                + "									<SchmeNm>\r\n"
                                + "										<Cd>"+ExCd+"</Cd>\r\n"
                                + "									</SchmeNm>\r\n"
                                + "								</Othr>\r\n"
                                + "							</Prsn>\r\n"
                                + "						</ExctgPrsn>");

                    else if (colCode.trim().equals("670"))
                        writeXml("            <AddtlAttrbts>\r\n"
                                + "             <RskRdcgTx>"+ RskRdcgTx +"</RskRdcgTx>\r\n"
                                + "             <SctiesFincgTxInd>"+securitiesFinTrnIndicator+"</SctiesFincgTxInd>\r\n"
                                + "            </AddtlAttrbts>");


                    else if (colCode.trim().equals("730")) {
                        writeXml( "          </New>\r\n"
                                + "        </Tx>");
                    }
                }


            }


        }

    }


    public static void writeXmlFI601()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String xsd;
        int cnt = 0;

        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);


        writeXml("<fi601 xmlns=\"http://www.bnr.ro/annex4\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://www.bnr.ro/annex4 FI601.xsd \">");



        if (currentSheetName.equals("Transactions")) {
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;

            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;


                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;
                int counter = 1;


                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals("") && colCode.equals("10"))
                        return;
                    if(cnt == 0) {
                        writeXml("			<transaction>");
                        cnt++;
                    }

                    if (colCode.trim().equals("10"))
                        writeXml("					<no>" + cellValue + "</no>");
                    else if (colCode.trim().equals("20"))
                        writeXml("					<code>" + cellValue + "</code>");
                    else if (colCode.trim().equals("30"))
                        writeXml("					<opCode>" + cellValue + "</opCode>");
                    else if (colCode.trim().equals("40"))
                        writeXml("					<time>" + cellValue + "</time>");
                    else if (colCode.trim().equals("50"))
                        writeXml("					<tradeDate>" + cellValue + "</tradeDate>");
                    else if (colCode.trim().equals("60"))
                        writeXml("					<settlementDate>" + cellValue + "</settlementDate>");
                    else if (colCode.trim().equals("70"))
                        writeXml("					<ISIN>" + cellValue + "</ISIN>");
                    else if (colCode.trim().equals("80")) {
                        writeXml("					<seller>");
                        writeXml("						<bic>" + cellValue + "</bic>");
                    }
                    else if (colCode.trim().equals("90"))
                        writeXml("						<account>" + cellValue + "</account>");
                    else if (colCode.trim().equals("100")) {
                        writeXml("						<clientType>" + cellValue + "</clientType>");
                        writeXml("					</seller>");
                    }
                    else if (colCode.trim().equals("110")) {
                        writeXml("					<buyer>");
                        writeXml("						<bic>" + cellValue + "</bic>");
                    }
                    else if (colCode.trim().equals("120"))
                        writeXml("						<account>" + cellValue + "</account>");
                    else if (colCode.trim().equals("130")) {
                        writeXml("						<clientType>" + cellValue + "</clientType>");
                        writeXml("					</buyer>");
                    }
                    else if (colCode.trim().equals("140"))
                        writeXml("					<yield>" + cellValue + "</yield>");
                    else if (colCode.trim().equals("150"))
                        writeXml("					<cleanPrice>" + cellValue + "</cleanPrice>");
                    else if (colCode.trim().equals("160"))
                        writeXml("					<dirtyPrice>" + cellValue + "</dirtyPrice>");
                    else if (colCode.trim().equals("170"))
                        writeXml("					<settlementAmount>" + cellValue + "</settlementAmount>");
                    else if (colCode.trim().equals("180"))
                        writeXml("					<nominalValue>" + cellValue + "</nominalValue>");
                    else if (colCode.trim().equals("190"))
                        writeXml("					<currency>" + cellValue + "</currency>");
                    else if (colCode.trim().equals("200")) {
                        writeXml("					<contract>");
                        writeXml("					<contractType>" + cellValue + "</contractType>");}
                    else if (colCode.trim().equals("210")) {
                        writeXml("					<contractNumber>" + cellValue + "</contractNumber>");
                        writeXml("					</contract>");
                    }
                    else if (colCode.trim().equals("220"))
                        writeXml("					<confirmationType>" + cellValue + "</confirmationType>");
                    else if (colCode.trim().equals("230"))
                        writeXml("					<platform>" + cellValue + "</platform>");
                    else if (colCode.trim().equals("240"))
                        writeXml("					<comments>" + cellValue + "</comments>");

                }
                writeXml("			</transaction>");
                cnt = 0;
            }

            writeXml("</fi601>");
        }
    }

    public static void writeXmlFI401()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType;
        String xsd;
        int cnt = 0;

        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);


        writeXml("<Participant-ReconImport xmlns=\"http://eps.transfond.ro/2004/FIMessages\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        writeXml("	<OSN>");
        writeXml("			<interfaceName>PARTtoGSRS</interfaceName>");
        writeXml("			<refDate>" + refDate + "</refDate>" );
        writeXml("			<seq>"+ String.valueOf(exId) +"</seq>");
        writeXml("	</OSN>");
        writeXml("	<statementDate>" + refDate + "</statementDate>");
        writeXml("	<participantId>" + valueOfB271 + "</participantId>");


        if (currentSheetName.equals("Amounts")) {
            int rowCount = sheet.getLastRowNum();
            if (rowCount == 0) return;

            // Map ISIN -> List of details [securitiesAccount, totalNominalValue, totalNominalValuePledged, totalNominalValueBanned]
            Map<String, List<String[]>> isinMap = new LinkedHashMap<>();

            for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if (isIgnoredColor(cell_check))
                    continue;

                String currentISIN = "";
                String[] details = new String[4]; // 0: securitiesAccount, 1: totalNominalValue, 2: totalNominalValuePledged, 3: totalNominalValueBanned

                for (int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++) {
                    Cell cell = row.getCell(columnIndex);
                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();
                    String cellValue = getCellValue(cell).trim();

                    if (cellValue.equals("")) continue;

                    switch (colCode) {
                        case "30":
                            currentISIN = cellValue;
                            break;
                        case "40":
                            details[0] = cellValue;
                            break;
                        case "50":
                            details[1] = cellValue;
                            break;
                        case "60":
                            details[2] = cellValue;
                            break;
                        case "70":
                            details[3] = cellValue;
                            break;
                    }
                }

                if (!currentISIN.equals("")) {
                    isinMap.computeIfAbsent(currentISIN, k -> new ArrayList<>()).add(details);
                }
            }

            for (Map.Entry<String, List<String[]>> entry : isinMap.entrySet()) {
                String isin = entry.getKey();
                List<String[]> detailsList = entry.getValue();

                writeXml("			<issueReconInfo>");
                writeXml("					<ISIN>" + isin + "</ISIN>");

                for (String[] details : detailsList) {
                    if (details[0] != null)
                        writeXml("					<accountReconInfo>");
                    writeXml("						<securitiesAccount>" + details[0] + "</securitiesAccount>");
                    if (details[1] != null)
                        writeXml("						<totalNominalValue>" + details[1] + "</totalNominalValue>");
                    if (details[2] != null)
                        writeXml("						<totalNominalValuePledged>" + details[2] + "</totalNominalValuePledged>");
                    if (details[3] != null)
                        writeXml("						<totalNominalValueBanned>" + details[3] + "</totalNominalValueBanned>");
                    writeXml("					</accountReconInfo>");
                }

                writeXml("			</issueReconInfo>");
            }
        }


    }



    public static void writeXmlBOP180() throws SQLException {
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;
        String messageType = null;
        String xsd;
        int cnt = 0;

        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if (valueOfC007 == null || valueOfC007.trim().equals("")) {
            valueOfC007 = "1";
        }

        valueOfB271 = "";

        if (tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);

        if (currentSheetName.equals("B1") && cnt == 0) {
            messageType = "BOP180-01";
            xsd = "BOP180-01.xsd";
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xsi:schemaLocation=\"http://extranet.bnr.ro " + xsd + "\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            writeXml("	<Header>");
            writeXml("	  <SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("      <SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("      <MessageType>" + messageType + "</MessageType>");
            writeXml("      <RefDate>" + refDate + "</RefDate>");
            writeXml("      <SenderMessageId>000</SenderMessageId>");
            writeXml("      <OperationType>T</OperationType>");
            writeXml("	</Header>");
            writeXml("  <Body>");
            cnt++;
        }

        if (currentSheetName.equals("B1") || currentSheetName.equals("B2") || currentSheetName.equals("B3") || currentSheetName.equals("B4")  || currentSheetName.equals("B5") || currentSheetName.equals("B6"))  {
            int rowCount = sheet.getLastRowNum();
            if (rowCount == 0)
                return;

            for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null)
                    return;

                String cod = getCellValue(row.getCell(2)).trim();
                String moneda = getCellValue(row.getCell(3)).trim();
                String tara = getCellValue(row.getCell(4)).trim();
                String suma1 = getCellValue(row.getCell(5)).trim();
                String suma2 = getCellValue(row.getCell(6)).trim();
                if (cod.trim().equals(""))
                    return;

                if(!suma1.equals("0"))
                {
                    writeXml("        <Operation>");
                    writeXml("            <OperationId>" + counter + "</OperationId>");
                    writeXml("            <BalanceCode>" + cod + "</BalanceCode>");
                    writeXml("            <CurrencyCode>" + moneda + "</CurrencyCode>");
                    writeXml("            <CountryCode>" + tara + "</CountryCode>");
                    if (VerificariBalanceCode.coduriIncasariSauPlati.contains(cod.trim())) {
                        writeXml("            <AmountType>2</AmountType>");
                    }else if (VerificariBalanceCode.coduriDoarPlati.contains(cod.trim())) {
                        writeXml("        <AmountType>3</AmountType>");
                    }else if (VerificariBalanceCode.coduriDoarIncasari.contains(cod.trim())) {
                        writeXml("        <AmountType>2</AmountType>");
                    }else if (VerificariBalanceCode.coduriSolduri.contains(cod.trim())) {
                        writeXml("        <AmountType>1</AmountType>");
                    }
                    writeXml("            <Amount>" + suma1 + "</Amount>");
                    writeXml("        </Operation>");
                    counter++;
                }
                if(!suma2.equals("0"))

                {
                    writeXml("        <Operation>");
                    writeXml("            <OperationId>" + counter + "</OperationId>");
                    writeXml("            <BalanceCode>" + cod + "</BalanceCode>");
                    writeXml("            <CurrencyCode>" + moneda + "</CurrencyCode>");
                    writeXml("            <CountryCode>" + tara + "</CountryCode>");
                    if (VerificariBalanceCode.coduriIncasariSauPlati.contains(cod.trim())) {
                        writeXml("            <AmountType>3</AmountType>");
                    }
                    else if (VerificariBalanceCode.coduriSolduri.contains(cod.trim())) {
                        writeXml("        <AmountType>8</AmountType>");
                    }else if (VerificariBalanceCode.coduriDoarPlati.contains(cod.trim())) {
                        writeXml("        <AmountType>3</AmountType>");
                    }
                    writeXml("            <Amount>" + suma2 + "</Amount>");
                    writeXml("        </Operation>");
                    counter++;
                }
            }
        }/*if (currentSheetName.equals("B5")) {
	        int rowCount = sheet.getLastRowNum();
	        if (rowCount == 0)
	            return;

	        for (int rowIndex = startRow; rowIndex <= rowCount; rowIndex++) {
	            Row row = sheet.getRow(rowIndex);
	            if (row == null)
	                return;

	            String cod = getCellValue(row.getCell(2)).trim();
	            String moneda = getCellValue(row.getCell(3)).trim();
	            String tara = getCellValue(row.getCell(4)).trim();
	            String suma1 = getCellValue(row.getCell(5)).trim();
	            String suma2 = getCellValue(row.getCell(6)).trim();
	            if (cod.trim().equals(""))
	            	return;

	            if(!suma1.equals("0"))
	            {
		            writeXml("        <Operation>");
	            	writeXml("            <OperationId>" + counter + "</OperationId>");
	                writeXml("            <BalanceCode>" + cod + "</BalanceCode>");
	                writeXml("            <CurrencyCode>" + moneda + "</CurrencyCode>");
	                writeXml("            <CountryCode>" + tara + "</CountryCode>");
	                if (VerificariBalanceCode.coduriIncasariSauPlati.contains(cod.trim())) {
	            	writeXml("            <AmountType>2</AmountType>");
	                }else if (VerificariBalanceCode.coduriDoarPlati.contains(cod.trim())) {
	                	writeXml("        <AmountType>3</AmountType>");
	                }else if (VerificariBalanceCode.coduriDoarIncasari.contains(cod.trim())) {
	                	writeXml("        <AmountType>2</AmountType>");
	                }else if (VerificariBalanceCode.coduriSolduri.contains(cod.trim())) {
	                	writeXml("        <AmountType>1</AmountType>");
	                }
	            	writeXml("            <Amount>" + suma1 + "</Amount>");
	            	writeXml("        </Operation>");
	                counter++;
	            }
	            if(!suma2.equals("0"))

	            {
		            writeXml("        <Operation>");
	            	writeXml("            <OperationId>" + counter + "</OperationId>");
	                writeXml("            <BalanceCode>" + cod + "</BalanceCode>");
	                writeXml("            <CurrencyCode>" + moneda + "</CurrencyCode>");
	                writeXml("            <CountryCode>" + tara + "</CountryCode>");
	                if (VerificariBalanceCode.coduriIncasariSauPlati.contains(cod.trim())) {
		            	writeXml("            <AmountType>3</AmountType>");
		            	}
	                else if (VerificariBalanceCode.coduriSolduri.contains(cod.trim())) {
	                	writeXml("        <AmountType>8</AmountType>");
	                }
	            	writeXml("            <Amount>" + suma2 + "</Amount>");
	            	writeXml("        </Operation>");
	                counter++;
	            }
	        }
	    }*/
    }



    public static void writeXmlSTM171()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;


        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);

        if (currentSheetName.equals("anexa1"))
        {
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml(
                    "<Message xmlns=\"http://extranet.bnr.ro\" xsi:schemaLocation=\"http://extranet.bnr.ro STM171-05.xsd\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date())
                    + "</SendingDate>");
            writeXml("		<MessageType>" + "STM171-05" + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "000" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
            writeXml("		<IC>");
            writeXml("			<A1>");


            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;


            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {

                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;
                int cntr = 0;

                Cell cell_check = row.getCell(stringToInt.get(startColumn));
                if(isIgnoredColor(cell_check))
                    continue;

                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);



                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);


                    if (cellValue.trim().equals(""))
                        continue;


                    if(cntr == 0) {
                        writeXml("				<Item>");
                        cntr++;
                        hasitemsA1 = true;
                    }

                    if (cellValue.trim().equals("-") || cellValue.trim().equals(""))
                        cellValue = "0";

                    if(colCode.trim().equals("0"))
                        writeXml("					<Code>" + cellValue + "</Code>");
                    else if (colCode.trim().equals("B"))
                        writeXml("					<Maturity>" + cellValue + "</Maturity>");
                    else if (colCode.trim().equals("C"))
                        writeXml("					<Country>" + cellValue + "</Country>");
                    else if (colCode.trim().equals("D"))
                        writeXml("					<Sector>" + cellValue + "</Sector>");
                    else if (colCode.trim().equals("E"))
                        writeXml("					<Currency>" + cellValue + "</Currency>");
                    else if (colCode.trim().equals("F"))
                        writeXml("					<Value1>" + cellValue + "</Value1>");
                    else if (colCode.trim().equals("G")) {
                        writeXml("					<Value2>" + cellValue + "</Value2>");
                        writeXml("				</Item>");
                    };
                }



            }


        }
        if (currentSheetName.equals("anexa1a"))
        {
            if (hasitemsA1) {
                writeXml("			</A1>");
            }
            int sht = 0;
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 11) {
                return;
            }

            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);
                if(row == null) {
                    return;

                }
                int cntr = 0;
                if(!getCellValue(row.getCell(1)).equals("") && sht == 0) {
                    writeXml("			<A1A>");
                    sht++;
                }


                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);

                    if (cellValue.trim().equals(""))
                        return;

                    if(cntr == 0) {
                        writeXml("				<Item>");
                        cntr++;
                        hasitemsA1a = true;
                    }

                    if(colCode.trim().equals("0"))
                        writeXml("					<Code>" + cellValue + "</Code>");
                    else if (colCode.trim().equals("B"))
                        writeXml("					<Maturity>" + cellValue + "</Maturity>");
                    else if (colCode.trim().equals("C"))
                        writeXml("					<Country>" + cellValue + "</Country>");
                    else if (colCode.trim().equals("D"))
                        writeXml("					<Sector>" + cellValue + "</Sector>");
                    else if (colCode.trim().equals("E"))
                        writeXml("					<Currency>" + cellValue + "</Currency>");
                    else if (colCode.trim().equals("F"))
                        writeXml("					<Value1>" + cellValue + "</Value1>");
                    else if (colCode.trim().equals("G")) {
                        writeXml("					<Value2>" + cellValue + "</Value2>");
                        writeXml("				</Item>");
                    };

                }


            }

        }

        if (currentSheetName.equals("anexa2"))
        {
            if (hasitemsA1a) {
                writeXml("			</A1A>");
            }
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;
            int sht = 0;


            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;
                int cntr = 0;

                if(!getCellValue(row.getCell(1)).equals("") && sht == 0) {
                    writeXml("			<A2>");
                    sht++;
                }
                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);
                    if (cellValue.trim().equals(""))
                        continue;

                    if(cntr == 0) {
                        writeXml("				<Item>");
                        cntr++;
                        hasitemsA2 = true;
                    }

                    if(colCode.trim().equals("0"))
                        writeXml("					<Code>" + cellValue + "</Code>");
                    else if (colCode.trim().equals("B"))
                        writeXml("					<Maturity>" + cellValue + "</Maturity>");
                    else if (colCode.trim().equals("C"))
                        writeXml("					<Country>" + cellValue + "</Country>");
                    else if (colCode.trim().equals("D"))
                        writeXml("					<Sector>" + cellValue + "</Sector>");
                    else if (colCode.trim().equals("E"))
                        writeXml("					<Currency>" + cellValue + "</Currency>");
                    else if (colCode.trim().equals("F")) {
                        writeXml("					<Value1>" + cellValue + "</Value1>");
                        writeXml("				</Item>");}
                }



            }


        }
        if (currentSheetName.equals("anexa3"))
        {
            if (hasitemsA2) {
                writeXml("			</A2>");
            }

            int sht = 0;
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;

            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;
                int cntr = 0;
                if(!getCellValue(row.getCell(1)).equals("") && sht == 0) {
                    writeXml("			<A3>");
                    sht++;
                }
                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);
                    if (cellValue.trim().equals(""))
                        return;

                    if(cntr == 0) {
                        writeXml("				<Item>");
                        cntr++;
                        hasitemsA3 = true;
                    }

                    if (cellValue.trim().equals(""))
                        cellValue = "0";
                    if(colCode.trim().equals("0"))
                        writeXml("					<Code>" + cellValue + "</Code>");
                    else if (colCode.trim().equals("B"))
                        writeXml("					<Value1>" + cellValue + "</Value1>");
                    else if (colCode.trim().equals("C")) {
                        writeXml("					<Value2>" + cellValue + "</Value2>");
                        writeXml("				</Item>");}
                }



            }


        }
        if (currentSheetName.equals("anexa4"))
        {
            if (hasitemsA3) {
                writeXml("			</A3>");
            }
            int sht = 0;
            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 11)
                return;


            for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);
                if(row == null)
                    return;
                int cntr = 0;
                if(!getCellValue(row.getCell(1)).equals("") && sht == 0) {
                    writeXml("			<A4>");
                    sht++;
                }
                for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                {
                    Cell cell = row.getCell(columnIndex);

                    if(isIgnoredColor(cell))
                        continue;

                    colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                    String cellValue = getCellValue(cell);
                    if (cellValue.trim().equals(""))
                        return;

                    if (cellValue.trim().equals(""))
                        return;

                    if(cntr == 0) {
                        writeXml("				<Item>");
                        cntr++;
                        hasitemsA4 = true;
                    }

                    if(colCode.trim().equals("0"))
                        writeXml("					<Code>" + cellValue + "</Code>");
                    else if (colCode.trim().equals("B"))
                        writeXml("					<Maturity>" + cellValue + "</Maturity>");
                    else if (colCode.trim().equals("C"))
                        writeXml("					<Country>" + cellValue + "</Country>");
                    else if (colCode.trim().equals("D"))
                        writeXml("					<Sector>" + cellValue + "</Sector>");
                    else if (colCode.trim().equals("E"))
                        writeXml("					<CodeCC>" + cellValue + "</CodeCC>");
                    else if (colCode.trim().equals("F"))
                        writeXml("					<CountryCC>" + cellValue + "</CountryCC>");
                    else if (colCode.trim().equals("G"))
                        writeXml("					<Currency>" + cellValue + "</Currency>");
                    else if (colCode.trim().equals("H"))
                        writeXml("					<Value1>" + cellValue + "</Value1>");
                    else if (colCode.trim().equals("I")) {
                        writeXml("					<Value2>" + cellValue + "</Value2>");
                        writeXml("				</Item>");}
                }


            }


        }
        if (hasitemsA4) {
            writeXml("			</A4>");
        }



    }


    public static void writeXmlSpm130()throws SQLException
    {
        //preparedStmt = connection.prepareStatement("insert into SPD270_values values (?, ?, ?, ?, ? ,?, ?, ?)");
        String variabilaUnitateBancara = reportingEntity;
        String rowCode;
        String colCode;


        String sheetName = sheet.getSheetName();

        valueOfC007 = getCellValue(sheet, tableC007Cell);
        if(valueOfC007 == null || valueOfC007.trim().equals(""))
            valueOfC007 = "1";
        valueOfB271 = "";

        if(tableB271Cell != null)
            valueOfB271 = getCellValue(sheet, tableB271Cell);
        double value1 = Double.parseDouble(valueOfB271);
        long rounded1 = Math.round(value1);
        int result1 = (int) rounded1;

        if (currentSheetName.equals("RON"))
        {
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml(
                    "<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro SPM130-03.xsd\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date())
                    + "</SendingDate>");
            writeXml("		<MessageType>" + "SPM130-03" + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "000" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
            writeXml("    <Appendix>");
            writeXml("      <UB_Code>" + reportingEntity.replace("i", "") + "</UB_Code>");

        }
        writeXml("			<ReserveCurrency>");
        writeXml("			    <CurrencyCode>" + sheetName + "</CurrencyCode>");
        writeXml("				<ExchangeRate>" + valueOfC007 + "</ExchangeRate>");
        writeXml("				<PredictedReserveCurrency>"+ result1 +"</PredictedReserveCurrency>");


        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);
            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

            if (rowCode == null || rowCode.trim().equals(""))
                continue;


            writeXml("				<SPM" + rowCode + ">");

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                Cell cell = row.getCell(columnIndex);

                if(isIgnoredColor(cell))
                    continue;

                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                String cellValue = getCellValue(cell);



                if (cellValue.trim().equals(""))
                    cellValue = "0";
                double value = Double.parseDouble(cellValue);
                long rounded = Math.round(value);
                int result = (int) rounded;
                double decimalValue = Double.parseDouble(cellValue);
                int intValue = (int) (decimalValue * 100);
                if(colCode.trim().equals("Suma"))
                    writeXml("					<Achieved>" + cellValue + "</Achieved>");
                else if (colCode.trim().equals("Rata RMO"))

                    writeXml("					<ReserveRate>" + intValue + "</ReserveRate>");
                else if (colCode.trim().equals("RMO"))
                    writeXml("					<Predicted>" + result + "</Predicted>");

            }

            writeXml("				</SPM" + rowCode + ">");


        }

        writeXml("			</ReserveCurrency>");
        //preparedStmt.executeBatch();

    }
    public static void writeXmlCESOP()throws SQLException
    {
        Date formattedDate;
        Timestamp formattedTimestamp;
        SimpleDateFormat dateFormatGaranti = new SimpleDateFormat("yyyy-MM-dd-hh.mm.ss.SSSSSS");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'+03:00'");
        HashMap<String, Integer> transMap = new HashMap<>();

        //for (int sheetIndex = 0; sheetIndex <= workbook.getNumberOfSheets(); sheetIndex++)
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
        {

            //String prev_TRANSACTION_IDENTIFIER = "";
            String prev_DOC_REF_ID = "";
            String prev_PAYEE_ID = "";
            String prev_DOC_TYPE_INDIC = "";


            Sheet sheet = workbook.getSheetAt(sheetIndex);

            if (!sheet.getSheetName().equals("Header") && !sheet.getSheetName().equals("Tranzactii"))
                continue;
            int rowCount = sheet.getLastRowNum();
            //if (sheetIndex == 0)
            if (sheet.getSheetName().equals("Header"))
            {
                Row headerRow = sheet.getRow(0); // First row contains headers
                HashMap<String, Integer> headerMap = new HashMap<>();

                for (Cell cell : headerRow) {
                    headerMap.put(getCellValue(cell), cell.getColumnIndex());
                }
                Row row = sheet.getRow(1);
                Timestamp timestamp = new Timestamp(System.currentTimeMillis());
                String MessageRefId =  getCellValue(row.getCell(headerMap.get("MESSAGE_REF_ID")));

                if (MessageRefId == null || MessageRefId.equals(""))
                    MessageRefId = UUID.randomUUID().toString();

                writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                writeXml("<CESOP xmlns=\"urn:ec.europa.eu:taxud:fiscalis:cesop:v1\" version=\"4.03\" xmlns:cm=\"urn:eu:taxud:commontypes:v1\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
                writeXml("  <MessageSpec>");
                writeXml("    <TransmittingCountry>" + getCellValue(row.getCell(headerMap.get("TRANSMITTING_COUNTRY"))) + "</TransmittingCountry>");
                writeXml("    <MessageType>"+ getCellValue(row.getCell(headerMap.get("MESSAGE_TYPE"))) +"</MessageType>");
                writeXml("    <MessageTypeIndic>"+ getCellValue(row.getCell(headerMap.get("MESSAGE_TYPE_INDIC"))) +"</MessageTypeIndic>");
                writeXml("    <MessageRefId>" + MessageRefId + "</MessageRefId>");
                writeXml("    <SendingPSP>");
                if (getCellValue(row.getCell(headerMap.get("PSP_ID_TYPE"))).equals("Other")) {
                    writeXml("      <PSPId PSPIdType=\"" + getCellValue(row.getCell(headerMap.get("PSP_ID_TYPE"))) + "\" PSPIdOther=\"" + getCellValue(row.getCell(headerMap.get("PSP_ID_OTHER"))) +"\">" + getCellValue(row.getCell(headerMap.get("PSP_ID"))) + "</PSPId>");
                }else {
                    writeXml("      <PSPId PSPIdType=\"" + getCellValue(row.getCell(headerMap.get("PSP_ID_TYPE"))) + "\">" + getCellValue(row.getCell(headerMap.get("PSP_ID"))) + "</PSPId>");
                }

                writeXml("      <Name nameType=\"" + getCellValue(row.getCell(headerMap.get("PSP_NAME_TYPE"))) + "\">" + getCellValue(row.getCell(headerMap.get("PSP_NAME"))) + "</Name>");
                writeXml("    </SendingPSP>");
                writeXml("    <ReportingPeriod>");
                writeXml("      <Quarter>" + getCellValue(row.getCell(headerMap.get("QUARTER"))) + "</Quarter>");
                writeXml("      <Year>" + getCellValue(row.getCell(headerMap.get("YEAR"))) + "</Year>");
                writeXml("    </ReportingPeriod>");
                writeXml("    <Timestamp>"+ sdf.format(timestamp) +"</Timestamp>");
                writeXml("  </MessageSpec>");

                writeXml("  <PaymentDataBody>");
                writeXml("    <ReportingPSP>");
                if (getCellValue(row.getCell(headerMap.get("PSP_ID_TYPE"))).equals("Other")) {
                    writeXml("      <PSPId PSPIdType=\"" + getCellValue(row.getCell(headerMap.get("PSP_ID_TYPE"))) + "\" PSPIdOther=\"" + getCellValue(row.getCell(headerMap.get("PSP_ID_OTHER"))) +"\">" + getCellValue(row.getCell(headerMap.get("PSP_ID"))) + "</PSPId>");
                }else {
                    writeXml("      <PSPId PSPIdType=\"" + getCellValue(row.getCell(headerMap.get("PSP_ID_TYPE"))) + "\">" + getCellValue(row.getCell(headerMap.get("PSP_ID"))) + "</PSPId>");
                }
                writeXml("    </ReportingPSP>");
                continue;
            }
            else if (sheet.getSheetName().equals("Tranzactii")) {
                Row colRow = sheet.getRow(0);
                int colNr = colRow.getLastCellNum();
                for(int columnIndex = 0; columnIndex <= colNr - 1; columnIndex++) {
                    Row transRow = sheet.getRow(0); // Assuming first row contains headers
                    Cell cell = transRow.getCell(columnIndex);
                    transMap.put(getCellValue(cell), columnIndex);
                }

                //System.out.println(rowCount);
                for(int rowIndex = 1; rowIndex <= rowCount+1; rowIndex++)
                {
                    if(rowIndex == rowCount + 1)
                    {

                        writeXml("      <DocSpec>");
                        writeXml("        <cm:DocTypeIndic>" + prev_DOC_TYPE_INDIC + "</cm:DocTypeIndic>");
                        writeXml("        <cm:DocRefId>" + prev_DOC_REF_ID + "</cm:DocRefId>");
                        writeXml("      </DocSpec>");

                        writeXml("    </ReportedPayee>");
                        continue; // sau break
                    }



                    Row transRow = sheet.getRow(rowIndex);


                    String REF_DATE = getCellValue(transRow.getCell(transMap.get("REF_DATE")));
                    if(REF_DATE.equals(""))
                        break;
                    String LOAD_VERSION = getCellValue(transRow.getCell(transMap.get("LOAD_VERSION")));
                    String SOURCE = getCellValue(transRow.getCell(transMap.get("SOURCE")));
                    String PAYEE_ID = getCellValue(transRow.getCell(transMap.get("PAYEE_ID")));
                    String COUNTRY = getCellValue(transRow.getCell(transMap.get("COUNTRY")));
                    String ACCOUNT_IDENTIFIER = getCellValue(transRow.getCell(transMap.get("ACCOUNT_IDENTIFIER")));
                    String COUNTRY_CODE = getCellValue(transRow.getCell(transMap.get("COUNTRY_CODE")));
                    String ACCOUNT_IDENTIFIER_TYPE = getCellValue(transRow.getCell(transMap.get("ACCOUNT_IDENTIFIER_TYPE")));
                    String REPRESENTATIVE_ID = getCellValue(transRow.getCell(transMap.get("REPRESENTATIVE_ID")));
                    String REPRESENTATIVE_NAME = getCellValue(transRow.getCell(transMap.get("REPRESENTATIVE_NAME")));
                    String REPRESENTATIVE_NAME_TYPE = getCellValue(transRow.getCell(transMap.get("REPRESENTATIVE_NAME_TYPE")));
                    String PSP_ID_TYPE = getCellValue(transRow.getCell(transMap.get("PSP_ID_TYPE")));
                    String DOC_TYPE_INDIC = getCellValue(transRow.getCell(transMap.get("DOC_TYPE_INDIC")));
                    String DOC_REF_ID = getCellValue(transRow.getCell(transMap.get("DOC_REF_ID")));
                    String CORR_DOC_REF_ID = getCellValue(transRow.getCell(transMap.get("CORR_DOC_REF_ID")));
                    String PAYEE_NAME = getCellValue(transRow.getCell(transMap.get("PAYEE_NAME")));
                    String PAYEE_NAME_TYPE = getCellValue(transRow.getCell(transMap.get("PAYEE_NAME_TYPE")));
                    String CATEGORY = getCellValue(transRow.getCell(transMap.get("CATEGORY")));
                    String VALUE = getCellValue(transRow.getCell(transMap.get("VALUE")));
                    String ISSUED_BY = getCellValue(transRow.getCell(transMap.get("ISSUED_BY")));
                    String TYPE = getCellValue(transRow.getCell(transMap.get("TYPE")));
                    String PAYEE_EMAIL_ADDRESS = getCellValue(transRow.getCell(transMap.get("PAYEE_EMAIL_ADDRESS")));
                    String PAYEE_WEB_PAGE = getCellValue(transRow.getCell(transMap.get("PAYEE_WEB_PAGE")));
                    String COUNTRY_CODE_ADDRESS = getCellValue(transRow.getCell(transMap.get("COUNTRY_CODE_ADDRESS")));
                    String STREET = getCellValue(transRow.getCell(transMap.get("STREET")));
                    String BUILDING_IDENTIFIER = getCellValue(transRow.getCell(transMap.get("BUILDING_IDENTIFIER")));
                    String SUITE_IDENTIFIER = getCellValue(transRow.getCell(transMap.get("SUITE_IDENTIFIER")));
                    String FLOOR_IDENTIFIER = getCellValue(transRow.getCell(transMap.get("FLOOR_IDENTIFIER")));
                    String DISTRICT_NAME = getCellValue(transRow.getCell(transMap.get("DISTRICT_NAME")));
                    String POB = getCellValue(transRow.getCell(transMap.get("POB")));
                    String POST_CODE = getCellValue(transRow.getCell(transMap.get("POST_CODE")));
                    String CITY = getCellValue(transRow.getCell(transMap.get("CITY")));
                    String COUNTRY_SUBENTITY = getCellValue(transRow.getCell(transMap.get("COUNTRY_SUBENTITY")));
                    String ADRESS_FREE = getCellValue(transRow.getCell(transMap.get("ADRESS_FREE")));
                    String LEGAL_ADDRESS_TYPE = getCellValue(transRow.getCell(transMap.get("LEGAL_ADDRESS_TYPE")));
                    String TRANSACTION_IDENTIFIER = getCellValue(transRow.getCell(transMap.get("TRANSACTION_IDENTIFIER")));
                    String IS_REFUND = getCellValue(transRow.getCell(transMap.get("IS_REFUND")));
                    if (IS_REFUND.toLowerCase().equals("n"))
                        IS_REFUND = "false";
                    else IS_REFUND = "true";
                    String CORR_TRANSACTION_IDENTIFIER = getCellValue(transRow.getCell(transMap.get("CORR_TRANSACTION_IDENTIFIER")));
                    String CURRENCY = getCellValue(transRow.getCell(transMap.get("CURRENCY")));
                    String AMOUNT = String.format("%.2f", Double.parseDouble(getCellValue(transRow.getCell(transMap.get("AMOUNT")))));
                    String PAYMENT_METHOD_TYPE = getCellValue(transRow.getCell(transMap.get("PAYMENT_METHOD_TYPE")));
                    String PAYMENT_METHOD_OTHER = getCellValue(transRow.getCell(transMap.get("PAYMENT_METHOD_OTHER")));
                    String INIT_AT_PHYSICAL_PREM_MERCHANT = getCellValue(transRow.getCell(transMap.get("INIT_AT_PHYSICAL_PREM_MERCHANT")));
                    if (INIT_AT_PHYSICAL_PREM_MERCHANT.toLowerCase().equals("n"))
                        INIT_AT_PHYSICAL_PREM_MERCHANT = "false";
                    else INIT_AT_PHYSICAL_PREM_MERCHANT = "true";
                    String PAYER_MS = getCellValue(transRow.getCell(transMap.get("PAYER_MS")));
                    String PAYER_MS_SOURCE = getCellValue(transRow.getCell(transMap.get("PAYER_MS_SOURCE")));
                    String PSP_ROLE_TYPE = getCellValue(transRow.getCell(transMap.get("PSP_ROLE_TYPE")));
                    String PSP_ROLE_OTHER = getCellValue(transRow.getCell(transMap.get("PSP_ROLE_OTHER")));
                    String DATETIME = getCellValue(transRow.getCell(transMap.get("DATETIME")));
                    String ACCOUNT_IDENTIFIER_OTHER = getCellValue(transRow.getCell(transMap.get("ACCOUNT_IDENTIFIER_OTHER")));
                    String REPRESENT_NAME_OTHER = getCellValue(transRow.getCell(transMap.get("REPRESENT_NAME_OTHER")));
                    String PAYEE_NAME_OTHER = getCellValue(transRow.getCell(transMap.get("PAYEE_NAME_OTHER")));
                    String TAX_ID_OTHER = getCellValue(transRow.getCell(transMap.get("TAX_ID_OTHER")));
                    String TRANSACTION_DATE_OTHER = getCellValue(transRow.getCell(transMap.get("TRANSACTION_DATE_OTHER")));


		    	/*try {

			    	formattedDate = dateFormatGaranti.parse(DATETIME);
			    	formattedTimestamp = new Timestamp(formattedDate.getTime());
			    	DATETIME = sdf.format(formattedTimestamp);
		    	}
		    	catch (Exception ex)
		    	{
		    		continue;
		    	}*/

                    String TRANSACTION_DATE_TYPE = getCellValue(transRow.getCell(transMap.get("TRANSACTION_DATE_TYPE")));



                    if(DOC_REF_ID.equals(""))
                        continue;

                    if((DOC_REF_ID.equals("#") && prev_DOC_REF_ID.equals("")) || (!prev_PAYEE_ID.equals(PAYEE_ID) && !prev_PAYEE_ID.equals("") && DOC_REF_ID.equals("#")))
                    {
                        DOC_REF_ID = UUID.randomUUID().toString();
                    }

                    if(!DOC_REF_ID.equals(prev_DOC_REF_ID) && (!prev_PAYEE_ID.equals(PAYEE_ID)))
                    {
                        if(!prev_DOC_REF_ID.equals(""))
                        {

                            writeXml("      <DocSpec>");
                            writeXml("        <cm:DocTypeIndic>" + prev_DOC_TYPE_INDIC + "</cm:DocTypeIndic>");
                            writeXml("        <cm:DocRefId>" + prev_DOC_REF_ID + "</cm:DocRefId>");
                            writeXml("      </DocSpec>");

                            writeXml("    </ReportedPayee>");
                        }
                        prev_DOC_REF_ID = DOC_REF_ID;
                        prev_PAYEE_ID = PAYEE_ID;
                        prev_DOC_TYPE_INDIC = DOC_TYPE_INDIC;

                        writeXml("    <ReportedPayee>");


                        if(PAYEE_NAME_TYPE.equals("Other")) {
                            writeXml("      <Name nameType=\""+ PAYEE_NAME_TYPE + "\" nameOther=\"" + PAYEE_NAME_OTHER + "\">" + PAYEE_NAME + "</Name>");
                        } else {
                            writeXml("<Name nameType=\""+ PAYEE_NAME_TYPE +"\">"+ PAYEE_NAME + "</Name>");
                        }
                        writeXml("      <Country>" + ISSUED_BY + "</Country>");
                        writeXml("      <Address legalAddressType=\"" + LEGAL_ADDRESS_TYPE + "\">");
                        writeXml("        <cm:CountryCode>" + COUNTRY_CODE_ADDRESS + "</cm:CountryCode>");
                        if (!STREET.equals("") || !COUNTRY_SUBENTITY.equals(""))
                        {
                            writeXml("        <cm:AddressFix>");
                            if(!STREET.equals(""))
                                writeXml("          <cm:Street>" + STREET + "</cm:Street>");
                            writeXml("          <cm:City>" + CITY + "</cm:City>");
                            if(!COUNTRY_SUBENTITY.equals(""))
                                writeXml("          <cm:CountrySubentity>" + COUNTRY_SUBENTITY + "</cm:CountrySubentity>");
                            writeXml("        </cm:AddressFix>");
                        }

                        writeXml("        <cm:AddressFree>" + ADRESS_FREE + "</cm:AddressFree>");
                        writeXml("      </Address>");

                        if(CATEGORY != null)
                        {
                            if(CATEGORY.equals("VATId"))
                            {
                                writeXml("      <TAXIdentification>");
                                writeXml("        <VATId issuedBy=\"" + ISSUED_BY + "\">" + VALUE + "</VATId>");
                                writeXml("      </TAXIdentification>");
                            }
                            else if (CATEGORY.equals("TAXId"))
                            {
                                writeXml("      <TAXIdentification>");
                                writeXml("        <TAXId issuedBy=\"" + ISSUED_BY + "\" type=\"" + TYPE + "\">" + VALUE + "</TAXId>");
                                if (TYPE.equals("Other")) {
                                    writeXml("        	TAXIdOther" + TAX_ID_OTHER + "</TAXIdOther>");
                                }
                                writeXml("      </TAXIdentification>");
                            }
                            else writeXml("      <TAXIdentification/>");
                        }
                        else writeXml("      <TAXIdentification/>");
                        if(ACCOUNT_IDENTIFIER_TYPE.equals("Other")) {
                            writeXml("      <AccountIdentifier CountryCode=\"" + COUNTRY_CODE + "\" type=\"" + ACCOUNT_IDENTIFIER_TYPE + "\"  accountIdentifierOther=\"" + ACCOUNT_IDENTIFIER_OTHER + "\">" + ACCOUNT_IDENTIFIER + "</AccountIdentifier>");
                        }else {
                            writeXml("      <AccountIdentifier CountryCode=\"" + COUNTRY_CODE + "\" type=\"" + ACCOUNT_IDENTIFIER_TYPE + "\">" + ACCOUNT_IDENTIFIER + "</AccountIdentifier>");
                        }
                    }
                    writeXml("      <ReportedTransaction IsRefund=\"" + IS_REFUND.toLowerCase() + "\">");
                    writeXml("        <TransactionIdentifier>" + TRANSACTION_IDENTIFIER + "</TransactionIdentifier>");
                    writeXml("        <DateTime transactionDateType=\"" + TRANSACTION_DATE_TYPE + "\">" + DATETIME + "</DateTime>");
                    if (TRANSACTION_DATE_TYPE.equals("Other")) {
                        writeXml("        	<transactionDateOther>" + TRANSACTION_DATE_OTHER +"</transactionDateOther>");
                    }
                    writeXml("        <Amount currency=\"" + CURRENCY + "\">" + AMOUNT + "</Amount>");
                    writeXml("        <PaymentMethod>");
                    if(PAYMENT_METHOD_TYPE.equals("Other")) {
                        writeXml("          <cm:PaymentMethodType=\"" + PAYMENT_METHOD_TYPE + "\" PaymentMethodOther=\"" + PAYMENT_METHOD_OTHER + "\">"+ "</cm:PaymentMethodType>");
                    }else {
                        writeXml("          <cm:PaymentMethodType>" + PAYMENT_METHOD_TYPE + "</cm:PaymentMethodType>");}
                    writeXml("        </PaymentMethod>");
                    writeXml("        <InitiatedAtPhysicalPremisesOfMerchant>" + INIT_AT_PHYSICAL_PREM_MERCHANT.toLowerCase() + "</InitiatedAtPhysicalPremisesOfMerchant>");
                    writeXml("        <PayerMS PayerMSSource=\"" + PAYER_MS_SOURCE + "\">" + PAYER_MS + "</PayerMS>");
                    writeXml("        <PSPRole>");
                    if (PSP_ROLE_TYPE.equals("Other")) {
                        writeXml("          <cm:PSPRoleType=\"" + PSP_ROLE_TYPE + "\" PSPRoleOther=\"" + PSP_ROLE_OTHER + "\">"+ "</cm:PSPRoleType>");
                    }else {
                        writeXml("          <cm:PSPRoleType>" + PSP_ROLE_TYPE + "</cm:PSPRoleType>");}
                    writeXml("        </PSPRole>");
                    writeXml("      </ReportedTransaction>");

                }
            }
        }
    }


    public static void writeXmlSpd279()throws SQLException
    {
        try {

            preparedStmt = connection.prepareStatement("insert into SPD279_values values (?, ?, ?, ?, ? ,?, ?, ?)");
            String code;
            String geo;
            String geoTerminal;
            String value1;
            String value2;
            String comment1;
            String comment2;
            //int nos = workbook.getNumberOfSheets();

            BigDecimal bigD;

            DecimalFormat df = new DecimalFormat("0.######");

            if (currentSheetName.trim().equals("tabel 1") && marker)
            {
                marker = false;
                connection.prepareStatement("delete from SPD279_values where execution_id = " + String.valueOf(exId)).execute();

                writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro SPD279-01.xsd\">");
                //writeXml("	<xsi:schemaLocation=\"http://extranet.bnr.ro SPD278-01.xsd\">");
                writeXml("	<Header>");
                writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
                writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
                writeXml("		<MessageType>" + "SPD279-01" + "</MessageType>");
                writeXml("		<RefDate>" + refDate + "</RefDate>");
                writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
                writeXml("		<OperationType>" + "T" + "</OperationType>");
                writeXml("	</Header>");
                writeXml("	<Body>");
                writeXml("		<Contacts>");





                for(int contactNo = 0; contactNo < nrContacte; contactNo++)
                {

                    //validari contacte
                    if(SPD279_Contacts[contactNo][1].replaceAll("\\s+", "").length() > 50)
                    {
                        writeXml("lungime " + SPD279_Contacts[contactNo][1].replaceAll("\\s+", "") + " > 50");
                        throw new Exception("lungime " + SPD279_Contacts[contactNo][1].replaceAll("\\s+", "") + " > 50");
                    }
                    if (SPD279_Contacts[contactNo][2].replaceAll("\\s+", "").length() > 30)
                    {
                        writeXml("lungime " + SPD279_Contacts[contactNo][2].replaceAll("\\s+", "") + " > 30");
                        throw new Exception("lungime " + SPD279_Contacts[contactNo][1].replaceAll("\\s+", "") + " > 30");
                    }
                    if (SPD279_Contacts[contactNo][3].replaceAll("\\s+", "").length() > 50)
                    {
                        writeXml("lungime " + SPD279_Contacts[contactNo][3].replaceAll("\\s+", "") + " > 50");
                        throw new Exception("lungime " + SPD279_Contacts[contactNo][3].replaceAll("\\s+", "") + " > 50");
                    }


                    writeXml("			<Contact>");

                    writeXml("				<Name>" + SPD279_Contacts[contactNo][1].replaceAll("\\s+", "") + "</Name>");
                    writeXml("				<Phone>" + SPD279_Contacts[contactNo][2] + "</Phone>");
                    writeXml("				<Email>" + SPD279_Contacts[contactNo][3] + "</Email>");

                    writeXml("			</Contact>");
                }

                writeXml("		</Contacts>");
            }


            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            if(rowCount == 0)
                return;

            //validare TableCode
            if(!currentSheetName.replace("tabel ", "").equals("1") && !currentSheetName.replace("tabel ", "").equals("2") && !currentSheetName.replace("tabel ", "").equals("3") && !currentSheetName.replace("tabel ", "").equals("4") && !currentSheetName.replace("tabel ", "").equals("5") && !currentSheetName.replace("tabel ", "").equals("6") && !currentSheetName.replace("tabel ", "").equals("7") && !currentSheetName.replace("tabel ", "").equals("8"))
            {
                writeXml("tabela " + currentSheetName.replace("tabel ", "") + " nu apartine de (1, 2, 3, 4, 5, 6, 7, 8)");
                throw new Exception("tabela " + currentSheetName.replace("tabel ", "") + " nu apartine de (1, 2, 3, 4, 5, 6, 7, 8)");
            }

            writeXml("		<TableGroup>");
            writeXml("			<TableCode>" + currentSheetName.replace("tabel ", "") + "</TableCode>");


            if (currentSheetName.trim().equals("tabel 4") || currentSheetName.trim().equals("tabel 5"))
            {

                for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);
                    if(row == null)
                        continue;
                    code = getCellValue(row.getCell(1));
                    geo = getCellValue(row.getCell(3));
                    geoTerminal = getCellValue(row.getCell(4));
                    value1 = getCellValue(row.getCell(5));
                    value2 = getCellValue(row.getCell(6));
                    comment1 = getCellValue(row.getCell(13));
                    comment2 = getCellValue(row.getCell(14));


                    if((value1 == null || value1.equals("") || value1.toLowerCase().equals("x"))
                            && (value2 == null || value2.equals("") || value2.toLowerCase().equals("x")))
                        continue;

                    //System.out.println(row.getCell(0));

                    //System.out.println(row.getCell(2));
                    //if(code.equals("COD") || code.equals("9.1") || code.equals("9.1.1.1") || code.equals("9.2") || code.equals("9.3") || code.equals("9.3.1.1") || code.equals("9.4") || code.equals("9.5"))	//nu inteleg 100% cum functioneaza asta, trebuie sa remediez schema mai tarziu
                    if(isIgnoredColor(row.getCell(1)))
                        //System.out.println(getFillColorHex(row.getCell(0)));
                        continue;

                    if(code.trim().equals("") || code.trim() == null)
                        continue;

                    //System.out.println(row.getCell(0));

                    /*
                    if(code.equals("4.1.2.3.1.1") && geo.equals("I"))
                    {

                    	comment1="bruh";
                    	//bd = new BigDecimal(value2);
                    	//System.out.println(new DecimalFormat("#.######").format(bd));
                    }
                    */

                    if(code.trim() != null && !code.trim().equals("") &&
                            (geo != null && !geo.equals("") || geoTerminal != null && !geoTerminal.equals("") || value1.trim() != null && !value1.trim().equals("") || comment1.trim() != null && !comment1.trim().equals("") || value2.trim() != null && !value2.trim().equals("") || comment2.trim() != null && !comment2.trim().equals(""))
                    )
                    {

                        //validari preliminare tag-uri
                        if(code.trim().length() > 20)
                        {
                            //writeXml("lungime " + code.trim() + " > 20");
                            throw new Exception("lungime " + code.trim() + " > 20");
                        }
                        if(geo != null && geo.toUpperCase().trim().length() > 7)
                        {
                            //writeXml("lungime " + geo.toUpperCase().trim() + " > 7");
                            throw new Exception("lungime " + geo.toUpperCase().trim() + " > 7");
                        }
                        if(geoTerminal != null && geoTerminal.trim().length() > 7)
                        {
                            //writeXml("lungime " + geoTerminal.trim() + " > 7");
                            throw new Exception("lungime " + geoTerminal.trim() + " > 7");
                        }
                        if(comment1 != null && comment1.trim().length() > 2)
                        {
                            //writeXml("lungime " + comment1.trim() + " > 2");
                            throw new Exception("lungime " + comment1.trim() + " > 2");
                        }
                        if(comment2 != null && comment2.trim().length() > 2)
                        {
                            //writeXml("lungime " + comment2.trim() + " > 2");
                            throw new Exception("lungime " + comment2.trim() + " > 2");
                        }

                        writeXml("			<Item>");

                        writeXml("				<Code>" + code.trim()  + "</Code>");
                        if(geo != null && !geo.equals(""))
                            writeXml("				<Geo>" + geo.toUpperCase().trim()  + "</Geo>");


                        if(geoTerminal != null && !geoTerminal.equals("") && !geoTerminal.equals("x") && !geoTerminal.equals("X"))
                            writeXml("				<GeoTerminal>" + geoTerminal.trim()  + "</GeoTerminal>");

                        if(value1.trim() != null && !value1.trim().equals("") && !value1.trim().equals("x") && !value1.trim().equals("X"))
                            writeXml("				<Value1>" + String.valueOf(Integer.parseInt(value1))  + "</Value1>");

                        if(comment1.trim() != null && !comment1.trim().equals("") && !comment1.trim().equals("x") && !comment1.trim().equals("X"))
                            writeXml("				<Comment1>" + comment1.trim()  + "</Comment1>");

                        if(value2.trim() != null && !value2.trim().equals("") && !value2.trim().equals("x") && !value2.trim().equals("X"))
                        {
                            bigD = new BigDecimal(value2);
                            writeXml("				<Value2>" + df.format(bigD)  + "</Value2>");
                        }

                        if(comment2.trim() != null && !comment2.trim().equals(""))
                            writeXml("				<Comment2>" + comment2.trim()  + "</Comment2>");


                        writeXml("			</Item>");


                        preparedStmt.setInt(1, exId);
                        preparedStmt.setString(2, code);
                        preparedStmt.setString(3, geo);
                        preparedStmt.setString(4, geoTerminal);
                        preparedStmt.setString(5, value1);
                        preparedStmt.setString(6, value2);
                        preparedStmt.setString(7, comment1);
                        preparedStmt.setString(8, comment2);
                        preparedStmt.addBatch();


                    }

                    // if (rowIndex == rowCount)
                    //writeXml("		</TableGroup>");

                }
            }
            else if(currentSheetName.trim().equals("tabel 2") || currentSheetName.trim().equals("tabel 3"))
            {

                for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);
                    if(row == null)
                        continue;
                    code = getCellValue(row.getCell(1));
                    geo = getCellValue(row.getCell(3));
                    value1 = getCellValue(row.getCell(4));
                    comment1 = getCellValue(row.getCell(8));


                    if(value1 == null || value1.equals("") || value1.toLowerCase().equals("x"))
                        continue;

                    //System.out.println(row.getCell(0));

                    //System.out.println(row.getCell(2));
                    //if(code.equals("COD") || code.equals("9.1") || code.equals("9.1.1.1") || code.equals("9.2") || code.equals("9.3") || code.equals("9.3.1.1") || code.equals("9.4") || code.equals("9.5"))	//nu inteleg 100% cum functioneaza asta, trebuie sa remediez schema mai tarziu
                    if(isIgnoredColor(row.getCell(1)))
                        //System.out.println(getFillColorHex(row.getCell(0)));
                        continue;

                    if(code.trim().equals("") || code.trim() == null)
                        continue;

                    //System.out.println(row.getCell(0));



                    //validari preliminare tag-uri
                    if(code.trim().length() > 20)
                    {
                        //writeXml("lungime " + code.trim() + " > 20");
                        throw new Exception("lungime " + code.trim() + " > 20");
                    }
                    if(geo != null && geo.toUpperCase().trim().length() > 7)
                    {
                        //writeXml("lungime " + geo.toUpperCase().trim() + " > 7");
                        throw new Exception("lungime " + geo.toUpperCase().trim() + " > 7");
                    }
                    if(comment1 != null && comment1.trim().length() > 2)
                    {
                        //writeXml("lungime " + comment1.trim() + " > 2");
                        throw new Exception("lungime " + comment1.trim() + " > 2");
                    }


                    if(code.trim() != null && !code.trim().equals("") &&
                            (geo != null && !geo.equals("") || value1.trim() != null && !value1.trim().equals("") || comment1.trim() != null && !comment1.trim().equals(""))
                    )
                    {


                        writeXml("			<Item>");

                        writeXml("				<Code>" + code.trim()  + "</Code>");
                        if(geo != null && !geo.equals(""))
                            writeXml("				<Geo>" + geo.toUpperCase().trim()  + "</Geo>");

                        if(value1.trim() != null && !value1.trim().equals("") && !value1.trim().equals("x") && !value1.trim().equals("X"))
                            writeXml("				<Value1>" + String.valueOf(Integer.parseInt(value1))  + "</Value1>");

                        if(comment1.trim() != null && !comment1.trim().equals("") && !comment1.trim().equals("x") && !comment1.trim().equals("X"))
                            writeXml("				<Comment1>" + comment1.trim()  + "</Comment1>");



                        writeXml("			</Item>");


                        preparedStmt.setInt(1, exId);
                        preparedStmt.setString(2, code);
                        preparedStmt.setString(3, geo);
                        preparedStmt.setString(4, null);
                        preparedStmt.setString(5, value1);
                        preparedStmt.setString(6, null);
                        preparedStmt.setString(7, comment1);
                        preparedStmt.setString(8, null);
                        preparedStmt.addBatch();

                    }

                    // if (rowIndex == rowCount)
                    // writeXml("		</TableGroup>");

                }
            }
            else if (currentSheetName.trim().equals("tabel 1"))
            {

                for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);
                    if(row == null)
                        continue;
                    code = getCellValue(row.getCell(1));
                    geo = getCellValue(row.getCell(3));
                    value1 = getCellValue(row.getCell(4));
                    value2 = getCellValue(row.getCell(5));
                    comment1 = getCellValue(row.getCell(12));
                    comment2 = getCellValue(row.getCell(13));



                    if((value1 == null || value1.equals("") || value1.toLowerCase().equals("x"))
                            && (value2 == null || value2.equals("") || value2.toLowerCase().equals("x")))
                        continue;

                    //System.out.println(row.getCell(0));

                    //System.out.println(row.getCell(2));
                    //if(code.equals("COD") || code.equals("9.1") || code.equals("9.1.1.1") || code.equals("9.2") || code.equals("9.3") || code.equals("9.3.1.1") || code.equals("9.4") || code.equals("9.5"))	//nu inteleg 100% cum functioneaza asta, trebuie sa remediez schema mai tarziu
                    if(isIgnoredColor(row.getCell(1)))
                        //System.out.println(getFillColorHex(row.getCell(0)));
                        continue;

                    if(code.trim().equals("") || code.trim() == null)
                        continue;

                    //System.out.println(row.getCell(0));





                    if(code.trim() != null && !code.trim().equals("") &&
                            (geo != null && !geo.equals("") && !value1.trim().equals("") || comment1.trim() != null && !comment1.trim().equals("") || value2.trim() != null && !value2.trim().equals("") || comment2.trim() != null && !comment2.trim().equals(""))
                    )
                    {


                        //validari preliminare tag-uri
                        if(code.trim().length() > 20)
                        {
                            //writeXml("lungime " + code.trim() + " > 20");
                            throw new Exception("lungime " + code.trim() + " > 20");
                        }
                        if(geo != null && geo.toUpperCase().trim().length() > 7)
                        {
                            //writeXml("lungime " + geo.toUpperCase().trim() + " > 7");
                            throw new Exception("lungime " + geo.toUpperCase().trim() + " > 7");
                        }
                        if(comment1 != null && comment1.trim().length() > 2)
                        {
                            //writeXml("lungime " + comment1.trim() + " > 2");
                            throw new Exception("lungime " + comment1.trim() + " > 2");
                        }
                        if(comment2 != null && comment2.trim().length() > 2)
                        {
                            //writeXml("lungime " + comment2.trim() + " > 2");
                            throw new Exception("lungime " + comment2.trim() + " > 2");
                        }

                        writeXml("			<Item>");

                        writeXml("				<Code>" + code.trim()  + "</Code>");
                        if(geo != null && !geo.equals(""))
                            writeXml("				<Geo>" + geo.toUpperCase().trim()  + "</Geo>");


                        if(value1.trim() != null && !value1.trim().equals("") && !value1.trim().equals("x") && !value1.trim().equals("X"))
                            writeXml("				<Value1>" + String.valueOf(Integer.parseInt(value1))  + "</Value1>");

                        if(comment1.trim() != null && !comment1.trim().equals("") && !comment1.trim().equals("x") && !comment1.trim().equals("X"))
                            writeXml("				<Comment1>" + comment1.trim()  + "</Comment1>");

                        if(value2.trim() != null && !value2.trim().equals("") && !value2.trim().equals("x") && !value2.trim().equals("X"))
                        {
                            bigD = new BigDecimal(value2);
                            writeXml("				<Value2>" + df.format(bigD)  + "</Value2>");
                        }

                        if(comment2.trim() != null && !comment2.trim().equals(""))
                            writeXml("				<Comment2>" + comment2.trim()  + "</Comment2>");


                        writeXml("			</Item>");


                        preparedStmt.setInt(1, exId);
                        preparedStmt.setString(2, code);
                        preparedStmt.setString(3, geo);
                        preparedStmt.setString(4, null);
                        preparedStmt.setString(5, value1);
                        preparedStmt.setString(6, value2);
                        preparedStmt.setString(7, comment1);
                        preparedStmt.setString(8, comment2);
                        preparedStmt.addBatch();
                        preparedStmt.clearParameters();

                    }

                    // if (rowIndex == rowCount)
                    //writeXml("		</TableGroup>");

                }
            }
            else
            {
                for(int rowIndex = startRow; rowIndex <= rowCount; rowIndex++)
                {
                    Row row = sheet.getRow(rowIndex);
                    if(row == null)
                        continue;
                    code = getCellValue(row.getCell(1));
                    geo = getCellValue(row.getCell(3));
                    value1 = getCellValue(row.getCell(4));
                    value2 = getCellValue(row.getCell(5));
                    comment1 = getCellValue(row.getCell(12));
                    comment2 = getCellValue(row.getCell(13));

                    if((value1 == null || value1.equals("") || value1.toLowerCase().equals("x"))
                            && (value2 == null || value2.equals("") || value2.toLowerCase().equals("x")))
                        continue;

                    //System.out.println(row.getCell(0));

                    //System.out.println(row.getCell(2));
                    //if(code.equals("COD") || code.equals("9.1") || code.equals("9.1.1.1") || code.equals("9.2") || code.equals("9.3") || code.equals("9.3.1.1") || code.equals("9.4") || code.equals("9.5"))	//nu inteleg 100% cum functioneaza asta, trebuie sa remediez schema mai tarziu
                    if(isIgnoredColor(row.getCell(1)))
                        //System.out.println(getFillColorHex(row.getCell(0)));
                        continue;

                    if(code.trim().equals("") || code.trim() == null)
                        continue;
                    //System.out.println(row.getCell(0));




                    if(code.trim() != null && !code.trim().equals("") &&
                            (geo != null && !geo.equals("") || value1.trim() != null && !value1.trim().equals("") || comment1.trim() != null && !comment1.trim().equals("") || value2.trim() != null && !value2.trim().equals("") || comment2.trim() != null && !comment2.trim().equals(""))
                    )
                    {


                        //validari preliminare tag-uri
                        if(code.trim().length() > 20)
                        {
                            //writeXml("lungime " + code.trim() + " > 20");
                            throw new Exception("lungime " + code.trim() + " > 20");
                        }
                        if(geo != null && geo.toUpperCase().trim().length() > 7)
                        {
                            //writeXml("lungime " + geo.toUpperCase().trim() + " > 7");
                            throw new Exception("lungime " + geo.toUpperCase().trim() + " > 7");
                        }
                        if(comment1 != null && comment1.trim().length() > 2)
                        {
                            //writeXml("lungime " + comment1.trim() + " > 2");
                            throw new Exception("lungime " + comment1.trim() + " > 2");
                        }
                        if(comment2 != null && comment2.trim().length() > 2)
                        {
                            //writeXml("lungime " + comment2.trim() + " > 2");
                            throw new Exception("lungime " + comment2.trim() + " > 2");
                        }

                        writeXml("			<Item>");

                        writeXml("				<Code>" + code.trim()  + "</Code>");
                        if(geo != null && !geo.equals("") && !geo.equals("x") && !geo.equals("X"))
                            writeXml("				<Geo>" + geo.toUpperCase().trim()  + "</Geo>");



                        if(value1.trim() != null && !value1.trim().equals("") && !value1.trim().equals("x") && !value1.trim().equals("X"))
                            writeXml("				<Value1>" + String.valueOf(Integer.parseInt(value1))  + "</Value1>");

                        if(comment1.trim() != null && !comment1.trim().equals("") && !comment1.trim().equals("x") && !comment1.trim().equals("X"))
                            writeXml("				<Comment1>" + comment1.trim()  + "</Comment1>");

                        if(value2.trim() != null && !value2.trim().equals("") && !value2.trim().equals("x") && !value2.trim().equals("X"))
                        {
                            bigD = new BigDecimal(value2);
                            writeXml("				<Value2>" + df.format(bigD)  + "</Value2>");
                        }

                        if(comment2.trim() != null && !comment2.trim().equals("") && !comment2.trim().equals("x") && !comment2.trim().equals("X"))
                            writeXml("				<Comment2>" + comment2.trim()  + "</Comment2>");


                        writeXml("			</Item>");

                        /*
                        preparedStmt.setInt(1, exId);
                        preparedStmt.setString(2, code);
                        preparedStmt.setString(3, geo);
                        preparedStmt.setString(4, geoTerminal);
                        preparedStmt.setString(5, value1);
                        preparedStmt.setString(6, value2);
                        preparedStmt.setString(7, comment1);
                        preparedStmt.setString(8, comment2);
                        preparedStmt.addBatch();
                        */
                    }

                    //if (rowIndex == rowCount)
                    //  writeXml("		</TableGroup>");

                }
            }


            writeXml("		</TableGroup>");
            preparedStmt.executeBatch();
        }
        catch(Exception e)
        {
            e.printStackTrace(new PrintWriter(sw));
            System.err.println("Exception at sheet: " + currentSheetName + "\n" + sw.toString());
            xmlWriter.close();
        }

    }


    //sheet-uri pentru ASF610
    public static void writeASF610()
    {
        DecimalFormat df = new DecimalFormat("0.##");
        String rowCode, colCode = "";

        if (sheet.getSheetName().trim().equals("Anexa1 - Companii") && valueOfI056.trim().equals("C01"))
        {
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro ASF610-01.xsd\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("		<MessageType>" + "ASF610-01" + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
        }

        if(valueOfI056.equals("C01"))
        {
            writeXml("    <Appendix>");
            writeXml("      <Code>1</Code>");
        }
        else if (valueOfI056.equals("P01"))
        {
            writeXml("    <Appendix>");
            writeXml("      <Code>2</Code>");
        }

        writeXml("      <Question>");
        writeXml("        <Code>" + valueOfI056.trim() + "</Code>");

        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            int valueNr = 0;
            Row row = sheet.getRow(rowIndex);

            rowCode = "";

            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();


            writeXml("        <Factor>");
            writeXml("          <Code>" + rowCode.trim() + "</Code>");

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                valueNr++;

                Cell cell ;

                try
                {
                    cell = row.getCell(columnIndex);
                }
                catch(Exception ex)
                {
                    continue;
                }


                if(isIgnoredColor(cell))
                    continue;

                colCode = "";

                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();


                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                    continue;


                writeXml("                <Value"+ valueNr + ">" + df.format(Double.parseDouble(cellValue))  +"</Value" + valueNr + ">");

            }

            writeXml("        </Factor>");

        }

        writeXml("      </Question>");

        if(valueOfI056.equals("C11") && refDate.compareTo("2024-12-31") <= refDate.compareTo(refDate))
        {
            writeXml("    </Appendix>");
        }
        else if (valueOfI056.equals("C12")) {
            writeXml("    </Appendix>");
        }
        else if (valueOfI056.equals("P16") && refDate.compareTo("2024-12-31") <= refDate.compareTo(refDate))
        {
            writeXml("    </Appendix>");
        }
        else if (valueOfI056.equals("P18")) {
            writeXml("    </Appendix>");
        }

    }

    //sheet-uri pentru RFC420
    public static void writeRFC420()
    {
        String rowCode, colCode = "";

        if(sheet.getSheetName().trim().equals("F3"))
        {
            rowCode="";
        }

        if (sheet.getSheetName().trim().equals("F1"))
        {
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro RFC420-01.xsd\">");
            //writeXml("	<xsi:schemaLocation=\"http://extranet.bnr.ro SPD278-01.xsd\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("		<MessageType>" + "RFC420-01" + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
        }

        writeXml("        <Form" + sheet.getSheetName().trim().substring(sheet.getSheetName().trim().length() - 1) + ">");

        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            int valueNr = 0;

            Row row = sheet.getRow(rowIndex);

            rowCode = "";

            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

            writeXml("            <Item>");
            writeXml("                <Code>" + rowCode +"</Code>");

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                valueNr++;

                Cell cell ;

                try
                {
                    cell = row.getCell(columnIndex);
                }
                catch(Exception ex)
                {
                    continue;
                }


                if(isIgnoredColor(cell))
                    continue;

                colCode = "";

                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();


                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                {
                    cellValue = "0";
                }


                writeXml("                <Value"+ valueNr + ">" + cellValue +"</Value" + valueNr + ">");


            }

            writeXml("            </Item>");
        }

        writeXml("        </Form" + sheet.getSheetName().trim().substring(sheet.getSheetName().trim().length() - 1) + ">");

    }

    //sheet-uri pentru RFC421 (identic cu 420 in afara de header
    public static void writeRFC421()
    {
        String rowCode, colCode = "";

        if(sheet.getSheetName().trim().equals("F3"))
        {
            rowCode="";
        }

        if (sheet.getSheetName().trim().equals("F1"))
        {
            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://extranet.bnr.ro RFC421-01.xsd\">");
            //writeXml("	<xsi:schemaLocation=\"http://extranet.bnr.ro SPD278-01.xsd\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity.replace("i", "") + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("		<MessageType>" + "RFC421-01" + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "12345" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
        }

        writeXml("        <Form" + sheet.getSheetName().trim().substring(sheet.getSheetName().trim().length() - 1) + ">");

        for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            int valueNr = 0;

            Row row = sheet.getRow(rowIndex);

            rowCode = "";

            rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

            if(rowCode.equals(""))
                continue;

            writeXml("            <Item>");
            writeXml("                <Code>" + rowCode +"</Code>");

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {
                valueNr++;

                Cell cell ;

                try
                {
                    cell = row.getCell(columnIndex);
                }
                catch(Exception ex)
                {
                    continue;
                }


                if(isIgnoredColor(cell))
                    continue;

                colCode = "";

                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();


                String cellValue = getCellValue(cell);

                if (cellValue.equals(""))
                {
                    cellValue = "0";
                }


                writeXml("                <Value"+ valueNr + ">" + cellValue +"</Value" + valueNr + ">");


            }

            writeXml("            </Item>");
        }

        writeXml("        </Form" + sheet.getSheetName().trim().substring(sheet.getSheetName().trim().length() - 1) + ">");

    }
    public static void writeXmlRpn650()
    {

        int counter = 0;
        boolean writeTagDetails = true;
        boolean isNewRow = true;
        boolean isFirstRow = true;
        boolean formHasData = false;
        String oldCode = "";
        String clientType = "";
        SimpleDateFormat refDateFormat = new SimpleDateFormat("yyyy-MM-dd");


        String xlsMapValue;

        Row row = null;

        try
        {
            row = sheet.getRow(startRow);
        }
        catch(Exception ex)
        {

        }

        if (currentSheetName.equals("Tabelul 1"))
        {

            String rpnReportName = "";
            try
            {

                //if( DateUtil.getJavaDate(Double.parseDouble(refDate)).before(refDateFormat.parse("31/12/2022")) )
                if(refDateFormat.parse(refDate).before(refDateFormat.parse("2022-12-31")))
                {
                    if(consolidatedFlag.equals("N"))
                        rpnReportName = "RPN650-03";
                    else
                        rpnReportName = "RPN651-03";
                }
                else
                {
                    if(consolidatedFlag.equals("N"))
                        rpnReportName = "RPN650-04";
                    else
                        rpnReportName = "RPN651-04";
                }
            }
            catch(Exception ex)
            {

            }

            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
            writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro " + rpnReportName + ".xsd" + "\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("		<MessageType>" + rpnReportName + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "1" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
            //writeXml("		<Type>C</Type>");
            writeXml("		<OwnFunds>"+ getCellValue(sheet, tableI056Cell) +"</OwnFunds>");
            writeXml("		<OwnFundsP>"+ getCellValue(sheet, tableB790Column) +"</OwnFundsP>");
            writeXml("		<Limit1>"+ getCellValue(sheet, tableB015Value) +"</Limit1>");
            writeXml("		<Limit2>"+ getCellValue(sheet, tableB272Cell) +"</Limit2>");
            writeXml("		<Form1>");

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                counter++;

                if(getCellValue(row.getCell(stringToInt.get("A"))).equals("TOTAL"))
                {
                    writeXml("			<Total_Value15>" + getCellValue(row.getCell(stringToInt.get("A") + 22)) + "</Total_Value15>");
                    writeXml("			<Total_Value16>" + getCellValue(row.getCell(stringToInt.get("A") + 23)) + "</Total_Value16>");
                }
                else if(getCellValue(row.getCell(stringToInt.get("A"))).contains("TOTAL aferent entitatilor"))
                {
                    writeXml("				</Details>");
                    writeXml("			</Item>");
                    writeXml("			<Total_Type1>" + getCellValue(row.getCell(stringToInt.get(startColumn) + 22)) + "</Total_Type1>");
                }
                else
                {


                    try
                    {
                        clientType = getCellValue(row.getCell(0));
                    }
                    catch(Exception ex)
                    {
                        clientType = "PJ";
                    }

                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        String cellValue = getCellValue(cell);

                        if(cellValue.equals("69"))
                            cellValue = "0";

                        if(rpn650XmlMapping.containsKey((columnIndex - stringToInt.get(startColumn) + 1) +""))
                        {
                            xlsMapValue = rpn650XmlMapping.get((columnIndex - stringToInt.get(startColumn) + 1) +"");

                            if(xlsMapValue.equals("Code") && !cellValue.equals(oldCode))
                            {
                                oldCode = cellValue;
                                if(!isFirstRow)
                                {
                                    writeXml("				</Details>");
                                    writeXml("			</Item>");
                                }
                                writeXml("			<Item>");

                                writeTagDetails = true;
                                isNewRow = true;
                                isFirstRow = false;
                            }
                            else if(xlsMapValue.contains("Code") && cellValue.equals(oldCode))
                            {
                                isNewRow = false;
                                writeXml("					<Item>");
                            }


                            if(xlsMapValue.contains("Value") && writeTagDetails)
                            {
                                writeXml("				<Details>");
                                writeXml("					<Item>");
                                writeTagDetails = false;
                            }


                            if(xlsMapValue.contains("Value") && !cellValue.trim().equals(""))
                                writeXml("						<" + xlsMapValue + ">" + cellValue + "</" + xlsMapValue + ">");
                            else if(isNewRow && !xlsMapValue.contains("Value"))
                            {
                                if(xlsMapValue.equals("DebQuantum"))
                                {
                                    if(clientType.equals("PJ") || clientType.equals("J"))
                                    {
                                        writeXml("				<" + xlsMapValue + ">" + cellValue + "</" + xlsMapValue + ">");
                                        //writeXml("				<DebFunction>" + "" + "</DebFunction>");
                                    }
                                    else
                                    {
                                        //writeXml("				<" + xlsMapValue + ">" + "" + "</" + xlsMapValue + ">");
                                        writeXml("				<DebFunction>" + cellValue + "</DebFunction>");
                                    }
                                }
                                else if (xlsMapValue.equals("CodeLEI"))
                                {
                                    if (!cellValue.trim().equals(""))
                                        writeXml("				<" + xlsMapValue + ">" + cellValue + "</" + xlsMapValue + ">");
                                }
                                else
                                    writeXml("				<" + xlsMapValue + ">" + cellValue + "</" + xlsMapValue + ">");
                            }
                        }

                    }

                    writeXml("					</Item>");
                }



                row = sheet.getRow(startRow + counter);
            }

            writeXml("		</Form1>");
        }
        else if (currentSheetName.equals("Tabelul 2"))
        {
            formHasData = false;

            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                if(formHasData == false)
                    writeXml("		<Form2>");

                formHasData = true;

                counter++;

                if(getCellValue(row.getCell(stringToInt.get(startColumn))).equals("TOTAL"))
                {
                    writeXml("			<Total>" + getCellValue(row.getCell(stringToInt.get(startColumn) + 3)) + "</Total>");
                }
                else
                {
                    writeXml("			<Item>");
                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        if(isIgnoredColor(cell))
                            continue;

                        String cellValue = getCellValue(cell);
                        //System.out.println(counter + " -> " + cellValue);
                        if(rpn650XmlMapping.containsKey((columnIndex - stringToInt.get(startColumn) + 1) +""))
                        {
                            xlsMapValue = rpn650XmlMapping.get((columnIndex - stringToInt.get(startColumn) + 1) +"");
                            writeXml("				<" + xlsMapValue + ">" + cellValue + "</" + xlsMapValue + ">");
                        }

                    }

                    writeXml("			</Item>");
                }



                row = sheet.getRow(startRow + counter);
            }

            if(formHasData == true)
                writeXml("		</Form2>");

        }
        else if (currentSheetName.equals("Tabelul 3"))
        {

            formHasData = false;


            while(row != null && checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
            {
                if(formHasData == false)
                    writeXml("		<Form3>");

                formHasData = true;

                counter++;

                if(getCellValue(row.getCell(stringToInt.get(startColumn))).equals("TOTAL"))
                {
                    writeXml("			<Total>" + getCellValue(row.getCell(stringToInt.get(startColumn) + 3)) + "</Total>");
                }
                else
                {
                    writeXml("			<Item>");
                    for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                    {
                        Cell cell = row.getCell(columnIndex);

                        if(isIgnoredColor(cell))
                            continue;

                        String cellValue = getCellValue(cell);

                        if(rpn650XmlMapping.containsKey((columnIndex - stringToInt.get(startColumn) + 1) +""))
                        {
                            xlsMapValue = rpn650XmlMapping.get((columnIndex - stringToInt.get(startColumn) + 1) +"");
                            writeXml("				<" + xlsMapValue + ">" + cellValue + "</" + xlsMapValue + ">");
                        }

                    }

                    writeXml("			</Item>");
                }


                row = sheet.getRow(startRow + counter);
            }
            if(formHasData == true)
                writeXml("		</Form3>");

        }
    }

    public static void writeXmlRpn640()
    {

        int counter = 0;
        boolean writeTagDetails = true;
        boolean isNewRow = true;
        boolean isFirstRow = true;
        boolean formHasData = false;
        String oldCode = "";
        String clientType = "";
        SimpleDateFormat refDateFormat = new SimpleDateFormat("yyyy-MM-dd");


        //sheet-uri cu breakdown pe valuta
        if(currentSheetName.startsWith("IRR_STD_POZ"))
        {
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++)
            {
                sheet = workbook.getSheetAt(sheetIndex);

                valueOfC007 = "";

                if (sheet.getSheetName().startsWith(currentSheetName) && currentSheetName.contains("IRR_STD_POZ"))
                {
                    if(sheetIndex == 0)
                    {
                        String rpnReportName = "";

                        if(consolidatedFlag.equals("N"))
                            rpnReportName = "RPN640-02";
                        else
                            rpnReportName = "RPN641-02";

                        writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                        writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
                        writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro " + rpnReportName + ".xsd" + "\">");
                        writeXml("	<Header>");
                        writeXml("		<SenderId>" + reportingEntity + "</SenderId>");
                        writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
                        writeXml("		<MessageType>" + rpnReportName + "</MessageType>");
                        writeXml("		<RefDate>" + refDate + "</RefDate>");
                        writeXml("		<SenderMessageId>" + "1" + "</SenderMessageId>");
                        writeXml("		<OperationType>" + "T" + "</OperationType>");
                        writeXml("	</Header>");
                        writeXml("	<Body>");
                        writeXml("    <Form1>");
                    }
                    valueOfC007 = getCellValue(sheet, tableC007Cell);

                    if(valueOfC007 == null || valueOfC007.equals(""))
                        continue;

                    if(valueOfC007.equals("ALTE VALUTE"))
                        valueOfC007 = "LEU";


                    for(int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {

                        Row row = sheet.getRow(rowIndex);

                        String rowCode = "";
                        rowCode = getCellValue(sheet, tableRowCodeColumn + (rowIndex + 1)).trim();

                        if(!checkIfRowHasValue(row, stringToInt.get(startColumn), stringToInt.get(endColumn)))
                            continue;

                        if(rowIndex == startRow)
                        {
                            writeXml("      <Currency>");
                            writeXml("        <CurrencyAcro>" + valueOfC007 + "</CurrencyAcro>");
                        }

                        if(rowIndex != endRow)	//ultimul rand e de total
                        {
                            writeXml("        <Item>");
                            writeXml("          <Code>" + rowCode + "</Code>");
                        }


                        for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
                        {
                            if(rowIndex != endRow)
                            {

                                Cell cell = row.getCell(columnIndex);

                                if(isIgnoredColor(cell))
                                    continue;

                                String colCode= "";

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                                String cellValue = getCellValue(cell);

                                if (colCode == null || cellValue.equals("") || colCode.equals(""))
                                    continue;

                                colCode = colCode.substring(1, 2); //extract 1 din (1)
                                writeXml("          <Value" + colCode + ">" + cellValue + "</Value" + colCode + ">");
                            }
                            else
                            {
                                Cell cell = row.getCell(columnIndex);

                                if(isIgnoredColor(cell))
                                    continue;

                                String colCode= "";

                                colCode = getCellValue(sheet, intToString.get(columnIndex) + tableColumnCodeRow).trim();

                                String cellValue = getCellValue(cell);

                                if (colCode == null || cellValue.equals("") || colCode.equals(""))
                                    continue;
                                if(columnIndex == stringToInt.get(endColumn) - 1)
                                    writeXml("        <Total1>" + cellValue + "</Total1>");
                                else if (columnIndex == stringToInt.get(endColumn))
                                    writeXml("        <Total2>" + cellValue + "</Total2>");

                            }


                        }
                        if(rowIndex != endRow)
                        {

                            writeXml("        </Item>");
                        }
                        else
                        {
                            writeXml("      </Currency>");
                        }
                    }

                }
            }

            writeXml("    </Form1>");
        }
        else	//IRR_STD_VAL
        {
            String fonduriProprii = valueOfC007;
            String valaoreAbsoluta = valueOfC200;
            String procentFP = valueOfC010;
            writeXml("    <Form2>");
            writeXml("      <Value1>" + fonduriProprii + "</Value1>");
            writeXml("      <Value2>" + valaoreAbsoluta + "</Value2>");
            writeXml("      <Value3>" + procentFP + "</Value3>");
            writeXml("    </Form2>");
        }


    }

    public static void writeXmlRPS500()
    {

        int counter = 0;
        boolean writeTagDetails = true;
        boolean isNewRow = true;
        boolean isFirstRow = true;
        boolean formHasData = false;
        String oldCode = "";
        String clientType = "";
        SimpleDateFormat refDateFormat = new SimpleDateFormat("yyyy-MM-dd");


        String xlsMapValue;

        if (currentSheetName.equals("structura capitalului"))
        {


            writeXml("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writeXml("<Message xmlns=\"http://extranet.bnr.ro\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
            writeXml("	xsi:schemaLocation=\"http://extranet.bnr.ro " + "RPS500-01" + ".xsd" + "\">");
            writeXml("	<Header>");
            writeXml("		<SenderId>" + reportingEntity + "</SenderId>");
            writeXml("		<SendingDate>" + new SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()) + "</SendingDate>");
            writeXml("		<MessageType>" + "RPS500-01" + "</MessageType>");
            writeXml("		<RefDate>" + refDate + "</RefDate>");
            writeXml("		<SenderMessageId>" + "1" + "</SenderMessageId>");
            writeXml("		<OperationType>" + "T" + "</OperationType>");
            writeXml("	</Header>");
            writeXml("	<Body>");
            writeXml("	  <Appendix1>");

            valueOfC007 = getCellValue(sheet, tableC007Cell);

            if(valueOfC007.contains("RON"))
                valueOfC007 = "RON";

            writeXml("	    <Currency>" + valueOfC007 + "</Currency>");


            Row row = null;

            try
            {
                row = sheet.getRow(startRow);
            }
            catch(Exception ex)
            {

            }

            for(int columnIndex = stringToInt.get(startColumn); columnIndex <= stringToInt.get(endColumn); columnIndex++)
            {

                Cell cell = row.getCell(columnIndex);

                String cellValue = getCellValue(cell);


                if(rps500XmlMapping.containsKey((columnIndex - stringToInt.get(startColumn) + 1) +""))
                {
                    xlsMapValue = rps500XmlMapping.get((columnIndex - stringToInt.get(startColumn) + 1) +"");

                    writeXml("	    <"+xlsMapValue + ">"  + cellValue + "</" + xlsMapValue + ">");


                }
            }

            writeXml("	  </Appendix1>");

        }
    }

    public static void writeCRS()throws SQLException
    {

        String id_inreg;
        String prev_id_inreg = "";
        String persid;
        String prev_persid = "";
        String tip_declar;
        String tip_ctr;
        String tip_pers;
        String tip_detinator_cont;
        String tip_ctr_text;
        String tip_detinator_cont_text;
        String data_nastere;
        String oras_nastere;
        String oras_nastere_text;

        BigDecimal bigD;
        //int nos = workbook.getNumberOfSheets();

        for (int sheetIndex = 0; sheetIndex < 1; sheetIndex++)
        {

            Sheet sheet = workbook.getSheetAt(sheetIndex);

            if (sheetIndex == 0)
            {
                writeXml("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writeXml("<F3000 xmlns=\"mfp:anaf:dgti:f3000:declaratie:v2\" xsi:schemaLocation=\"http://extranet.bnr.ro F3000.xsd\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" luna=\"12\" an=\"2022\" cui=\"361897\" giin=\"###\" denumire=\"CEC BANK\" nume_declar=\"XXXXX\" prenume_declar=\"YYY\" functie_declar=\"Director\" tip_raportare=\"1\" tip_raportor=\"###\" totalPlata_A=\"3756\">");

            }

            int rowCount = sheet.getLastRowNum();
            //System.out.println(rowCount);
            for(int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                Row row = sheet.getRow(rowIndex);

                id_inreg = getCellValue(row.getCell(3));
                if(!id_inreg.equals(prev_id_inreg))
                {
                    writeXml("  <inregistrare id_inreg=\""+ id_inreg +"\">");
                }

                persid = getCellValue(row.getCell(4));
                if(!id_inreg.equals(prev_persid))
                {
                    tip_declar = getCellValue(row.getCell(5));
                    tip_ctr = getCellValue(row.getCell(6));
                    tip_pers = getCellValue(row.getCell(7));
                    tip_detinator_cont = getCellValue(row.getCell(9));
                    data_nastere = getCellValue(row.getCell(10));
                    oras_nastere = getCellValue(row.getCell(11));
                    if(tip_ctr.isEmpty())
                    {
                        tip_ctr_text = "";
                    }
                    else
                    {
                        tip_ctr_text = " tip_ctr=\"" + tip_ctr + "\"";
                    }
                    if(tip_detinator_cont.isEmpty())
                    {
                        tip_detinator_cont_text = "";
                    }
                    else
                    {
                        tip_detinator_cont_text = " tip_detinator_cont=\"" + tip_detinator_cont + "\"";
                    }
                    if(oras_nastere.isEmpty())
                    {
                        oras_nastere_text = "";
                    }
                    else
                    {
                        oras_nastere_text = " oras_nastere=\"" + oras_nastere + "\"";
                    }
                    writeXml("    <persoane tip_declar=\"" + tip_declar + "\"" + tip_ctr_text + " tip_pers+\"" + tip_pers + "\"" + tip_detinator_cont_text + " data_nastere=\"" + data_nastere + "\"" + oras_nastere_text
                            + ""
                    );
                }




            }
        }



        writeXml("</F3000>");
        //preparedStmt.executeBatch();

    }


    public static void main(String[] args)
    {
        try
        {
            if (args[0].startsWith("--") && args[0].contains("="))
            {

                Map<String, String> arguments = new HashMap<>();

                for (String arg : args) {
                    // Check if the argument is in key=value format
                    if (arg.startsWith("--") && arg.contains("=")) {
                        String[] keyValue = arg.substring(2).split("=", 2);
                        if (keyValue.length == 2) {
                            arguments.put(keyValue[0], keyValue[1]);
                        }
                    }
                }

                exId = Integer.valueOf(arguments.getOrDefault("exId", "Unknown"));
                reportsGroup = arguments.getOrDefault("exId", "Unknown"); //idk, legacy
                refDate = arguments.getOrDefault("refDate", "Unknown");
                consolidatedFlag = arguments.getOrDefault("consolidatedFlag", "Unknown");
                pathToJar = arguments.getOrDefault("pathToJar", "Unknown");
                xlsFileName = arguments.getOrDefault("xlsFileName", "Unknown");
                encryptKey = arguments.getOrDefault("encryptKey", "QUwSE9rc");
                connectionString = arguments.getOrDefault("connectionString", "Unknown");
                if (!connectionString.equals("Unknown"))
                {
                    isEncrypted = true;
                }
                else {
                    isEncrypted = false;
                }

            }
            else
            {
                isEncrypted = false;

                if(args.length == 5)
                {
                    if(isNumeric(args[0]))
                        exId = Integer.valueOf(args[0]);
                    reportsGroup = args[0];
                    refDate = args[1];
                    consolidatedFlag = args[2];
                    pathToJar = args[3];
                    xlsFileName = args[4];
                }
                else
                {
                    if(args.length == 6)
                    {
                        if(isNumeric(args[0]))
                            exId = Integer.valueOf(args[0]);
                        reportsGroup = args[0];
                        refDate = args[1];
                        consolidatedFlag = args[2];
                        pathToJar = args[3];
                        xlsFileName = args[4];
                        encryptKey = args[5];
                        isEncrypted = true;
                    }
                    else
                    {
                        String arguments = "";
                        for(String item : args)
                            arguments += item + " ";
                        throw new Exception("Invalid number of arguments, expected: 5 or 6 got: " + args.length + " -> Arguments: " + arguments);
                    }

                }
            }

            long startTime = System.currentTimeMillis();

            //initializez

            init();

            if((xlsFileId.equals("11")  && xlsFileName.startsWith("MACHETA")) || (xlsFileId.equals("12")  && xlsFileName.startsWith("SPD")) || (xlsFileId.equals("13")  && xlsFileName.startsWith("SPE")) || xlsFileId.equals("17") || xlsFileId.equals("20") || xlsFileId.equals("21") || xlsFileId.equals("22") || xlsFileId.equals("23") || xlsFileId.equals("25") || xlsFileId.equals("28") || xlsFileId.equals("29") || xlsFileId.equals("30")  && xlsFileName.startsWith("Macheta_CESOP") || xlsFileId.equals("34")  && xlsFileName.startsWith("SPM130") || xlsFileName.startsWith("PVL") || xlsFileName.startsWith("STM171")|| xlsFileName.startsWith("PMI120")|| xlsFileName.startsWith("STA-STV") || xlsFileName.startsWith("BOP180") || xlsFileName.startsWith(("FI601"))|| xlsFileName.startsWith(("FI401"))|| xlsFileName.startsWith(("ARM"))|| xlsFileName.startsWith(("STM231"))|| xlsFileName.startsWith(("MACHETA_RPS100"))) //Xml Bnr
            {
                if(xlsFileId.equals("12")  && xlsFileName.startsWith("SPD"))
                    writeXmlSpd278();
                else if(xlsFileId.equals("13")  && xlsFileName.startsWith("SPE"))
                    writeXmlSpe273();
                else if (xlsFileId.equals("20")  && xlsFileName.startsWith("SPD"))
                    writeXmlSpd270();
                else if (xlsFileId.equals("30")  && xlsFileName.startsWith("Macheta_CESOP"))
                    writeXmlCESOP();
                //else if (xlsFileId.equals("21")  && xlsFileName.startsWith("SPD"))
                //writeXmlSpd279();
            }
            else
            {
                writeXml("<?xml version='1.0' encoding='ISO-8859-1'?>");
                writeXml("<document name=\"" + documentName + "\" level=\"" + reportingEntity + "\" date=\"" + refDate + "\">");

                if(xlsFileId.equals("16"))
                {
                    writeXmlIMV("<?xml version='1.0' encoding='ISO-8859-1'?>");
                    writeXmlIMV("<document name=\"" + documentName + "\" level=\"" + reportingEntity + "\" date=\"" + refDate + "\">");
                    writeXmlRM("<?xml version='1.0' encoding='ISO-8859-1'?>");
                    writeXmlRM("<document name=\"" + documentName + "\" level=\"" + reportingEntity + "\" date=\"" + refDate + "\">");
                }
            }

            abacusCountryCodes.clear();

            if(xlsFileId.equals("2"))
            {
                sqlQuery = "select country,\r\n" +
                        "       iso_code,\r\n" +
                        "       c010\r\n" +
                        "  from rep_xls_country_code_mapping\r\n";

                resultSet = statement.executeQuery(sqlQuery);

                while(resultSet.next())
                {
                    abacusCountryCodes.put(resultSet.getString(1), resultSet.getString(3));
                    abacusCountryCodes.put(resultSet.getString(2), resultSet.getString(3));
                }
            }

            //trec prin toate sheet-urile din lista
            for (int i = 0; i < sheetIdList.size(); i++)
            {
                currentSheetName = sheetNameList.get(i);
                reportName = sheetReportNameMap.get(sheetIdList.get(i));//(currentSheetName);
                xmlRowNamePattern = sheetRowNameMap.get(sheetIdList.get(i));//(currentSheetName);
                xmlColumnNamePattern = sheetColumnNameMap.get(sheetIdList.get(i));//(currentSheetName);
                xmlPattern = sheetXmlPatternMap.get(sheetIdList.get(i));//(currentSheetName);

                abacusRowNames.clear();

                sqlQuery = "select abacus_row_code , " +
                        "       its_row_code\r\n" +
                        "  from rep_xls_row_code_mapping\r\n" +
                        " where sheet_id = " + (String)sheetIdList.get(i);

                resultSet = statement.executeQuery(sqlQuery);

                while(resultSet.next())
                {
                    abacusRowNames.put(resultSet.getString(2), resultSet.getString(1));
                }

                //RPN650
                if(xlsFileId.equals("17"))
                {
                    rpn650XmlMapping.clear();
                    sqlQuery = "select column_name , " +
                            "       position_in_xls\r\n" +
                            "  from rep_xls_rpn650_mapping\r\n" +
                            " where sheet_id = " + (String)sheetIdList.get(i);

                    resultSet = statement.executeQuery(sqlQuery);

                    while(resultSet.next())
                    {
                        rpn650XmlMapping.put(resultSet.getString(2), resultSet.getString(1));
                    }

                }

                if(xlsFileId.equals("29"))
                {
                    rps500XmlMapping.clear();
                    sqlQuery = "select column_name , " +
                            "       position_in_xls\r\n" +
                            "  from rep_xls_rps500_mapping\r\n" +
                            " where sheet_id = " + (String)sheetIdList.get(i);

                    resultSet = statement.executeQuery(sqlQuery);

                    while(resultSet.next())
                    {
                        rps500XmlMapping.put(resultSet.getString(2), resultSet.getString(1));
                    }

                }



                //trec prin toate tabelele definite pe sheet-ul respectiv
                sqlQuery = "select xt.table_range," + newLine +
                        "       xt.table_row_code_column," + newLine +
                        "       xt.table_column_code_row," + newLine +
                        "       xt.table_i056_cell," + newLine +
                        "       xt.table_b790_column," + newLine +
                        "       xt.table_b015_value," + newLine +
                        "       xt.table_b272_cell," + newLine +
                        "       xt.table_groupfield_column," + newLine +
                        "       xt.table_b271_cell," + newLine +
                        "       xt.table_rank_column," + newLine +
                        "       xt.table_c007_column," + newLine +
                        "       xt.table_c200_column," + newLine +
                        "       xt.table_c010_column" + newLine +
                        "  from rep_xls_sheet_tables xt" + newLine +
                        " where xt.sheet_id = " + sheetIdList.get(i) + newLine +
                        "   and " + refDateSql + " between xt.valid_from and xt.valid_to";

                resultSet = statement.executeQuery(sqlQuery);


                while(resultSet.next())
                {
                    tableRange = resultSet.getString(1);
                    tableRowCodeColumn = resultSet.getString(2);
                    tableColumnCodeRow = resultSet.getString(3);
                    tableI056Cell = resultSet.getString(4);
                    tableB790Column = resultSet.getString(5);
                    tableB015Value = resultSet.getString(6);
                    tableB272Cell = resultSet.getString(7);
                    tableGroupfieldColumn = resultSet.getString(8);
                    tableB271Cell = resultSet.getString(9);
                    tableRankColumn = resultSet.getString(10);
                    tableC007Cell = resultSet.getString(11);
                    tableC200Cell = resultSet.getString(12);
                    tableC010Cell = resultSet.getString(13);

                    startColumn = getStartColumn(tableRange);
                    endColumn = getEndColumn(tableRange);
                    startRow = getStartRow(tableRange);
                    endRow = getEndRow(tableRange);


                    //ma duc pe sheet-ul respectiv
                    sheet = workbook.getSheet(currentSheetName);

                    valueOfC007 = "";

                    if(tableC007Cell != null)
                        valueOfC007 = getCellValue(sheet, tableC007Cell);

                    valueOfC200 = "";

                    if(tableC200Cell != null)
                        valueOfC200 = getCellValue(sheet, tableC200Cell);

                    valueOfC010 = "";

                    if(tableC010Cell != null)
                        valueOfC010 = getCellValue(sheet, tableC010Cell);

                    valueOfI056 = "";

                    if(tableI056Cell != null)
                        valueOfI056 = getCellValue(sheet, tableI056Cell);

                    valueOfB272 = "";

                    if(tableB272Cell != null)
                        valueOfB272 = getCellValue(sheet, tableB272Cell);

                    //xlsFileName = xlsFileName.replaceAll("[ _]", "");
                    //cazul COREP - 6.2 - Group Solvency - dinamic + rand
                    if(xlsFileId.equals("1") && (new SimpleDateFormat("yyyy-MM-dd").parse(refDate)).after(refDateBnrFormat.parse("29/06/2021")))
                        continue;
                    else if (xlsFileId.equals("1") && currentSheetName.equals("6.2"))
                        writeXmlCorep6_2();
                        //cazul COREP - 8.2 - CR IRB - Obligor Pools - dinamic (nr variabil de randuri)
                    else if (xlsFileId.equals("1") && currentSheetName.startsWith("8.2"))
                        writeXmlCorep8_2();
                        //cazul COREP - 14 - Securitisation Details - dinamic + rand
                    else if (xlsFileId.equals("1") && (currentSheetName.equals("14") || currentSheetName.equals("14.1")))
                        writeXmlCorep14();
                        //cazul COREP - 10.2 - CR EQU IRB - Obligor Pools - dinamic (nr variabil de randuri)
                    else if (xlsFileId.equals("1") && currentSheetName.equals("10.2"))
                        writeXmlCorep10_2();
                        //cazul COREP - 17.2 - OPR DETAILS 2 - dinamic (nr variabil de randuri)
                    else if (xlsFileId.equals("1") && currentSheetName.equals("17.2"))
                        writeXmlCorep17_2();
                        //cazul COREP - 32.3 - PRUVAL 3 - dinamic (nr variabil de randuri)
                    else if (xlsFileId.equals("1") && currentSheetName.equals("32.3"))
                        writeXmlCorep32_3();
                        //cazul COREP - 32.4 - PRUVAL 4 - dinamic (nr variabil de randuri)
                    else if (xlsFileId.equals("1") && currentSheetName.equals("32.4"))
                        writeXmlCorep32_4();
                        //cazul COREP - 9.1 - CR GB 1 - Geographical Breakdown SA Exposures - sheet multiplicat dinamic
                        //cazul COREP - 9.2 - CR GB 2 - Geographical Breakdown IRB Exposures - sheet multiplicat dinamic
                        //cazul COREP - 9.4 - CR GB 4 - Geographical Breakdown countercyclical (CCB) - sheet multiplicat dinamic
                        //cazul COREP - 18 - MKR TDI - Traded Debt Instruments - sheet multiplicat dinamic
                        //cazul COREP - 21 - MKR EQU - Equities - sheet multiplicat dinamic
                        //cazul COREP - 33 - GENERAL GOV - Equities - sheet multiplicat dinamic
                    else if ((xlsFileId.equals("1") ) &&
                            (currentSheetName.equals("9.1") ||
                                    currentSheetName.equals("9.2") ||
                                    currentSheetName.equals("9.4") ||
                                    currentSheetName.equals("18") ||
                                    currentSheetName.equals("21") ||
                                    currentSheetName.equals("33")))
                        writeXmlCorepDynamicSheet();
                        //cazul IP_LOSSES - 15 - sheet multiplicat dinamic
                    else if ((xlsFileId.equals("3") || xlsFileId.equals("1")) &&
                            currentSheetName.startsWith("15"))
                        writeXmlCorepDynamicSheet();
                        //sheet-uri default COREP
                    else if (xlsFileId.equals("1") ||
                            xlsFileId.equals("3") ||
                            xlsFileId.equals("5"))
                        writeXmlCorepDefault();
                    else if((xlsFileId.equals("6") && xlsFileName.contains("CCY")) || xlsFileName.equals("Annex 12 (Liquidity) CCY.xlsx"))
                        writeXmlNsfrCcyDefault();
                    else if (xlsFileId.equals("6"))
                        writeXmlNsfrDefault();
                    else if((xlsFileId.equals("7") && xlsFileName.contains("CCY")) || xlsFileName.equals("Annex 24 (LCR) CCY.xlsx"))
                        writeXmlLcrCcyDefault();
                    else if(xlsFileId.equals("7"))
                        writeXmlLcrDefault();
                    else if((xlsFileId.equals("9") && xlsFileName.contains("CCY")) || xlsFileName.equals("Annex 18 (AMM) CCY.xlsx"))
                        writeXmlAlmmCcyDefault();
                    else if(xlsFileId.equals("9"))
                        writeXmlAlmmDefault();
                    else if(xlsFileId.equals("4"))
                        writeXmlLeDefault();
                    else if(xlsFileId.equals("8"))
                        writeXmlAeDefault();
                    else if(xlsFileId.equals("2"))
                        writeXmlFinrepDefault();
                    else if(xlsFileId.equals("10"))
                        writeXmlResolutionDefault();
                    else if(xlsFileId.equals("11") && xlsFileName.startsWith("MACHETA"))
                        writeXmlFinrepBnr();
                    else if(xlsFileId.equals("14"))
                        writeXmlCovid19();
                    else if(xlsFileId.equals("15"))
                        writeXmlFundingPlan();
                    else if(xlsFileId.equals("16"))
                        writeXmlSbp();
                    else if(xlsFileId.equals("17"))
                        writeXmlRpn650();
                    else if(xlsFileId.equals("21"))
                        writeXmlSpd279();
                    else if(xlsFileId.equals("22") && !reportingEntity.contains("317"))
                        writeASF610();
                    else if(xlsFileId.equals("25") && reportingEntity.contains("317"))
                        writeASF610();
                    else if(xlsFileId.equals("23") && xlsFileName.equals("RFC420.xlsx"))
                        writeRFC420();
                    else if(xlsFileId.equals("23") && xlsFileName.equals("RFC421.xlsx"))
                        writeRFC421();
                    else if(xlsFileId.equals("23") && xlsFileName.equals("RFC421.xlsx"))
                        writeRFC421();
                    else if(xlsFileId.equals("24"))
                        writeCRS();
                    else if(xlsFileId.equals("28") && (xlsFileName.equals("Macheta_RPN640.xlsx") || xlsFileName.equals("Macheta_RPN641.xlsx")))
                        writeXmlRpn640();
                    else if (xlsFileId.equals("29"))
                        writeXmlRPS500();
                    else if (xlsFileId.equals("34")  && xlsFileName.startsWith("SPM130"))
                        writeXmlSpm130();
                    else if (xlsFileId.equals("37")  && xlsFileName.startsWith("STM171"))
                        writeXmlSTM171();
                    else if (xlsFileName.startsWith("PVL110"))
                        writeXmlPVL110();
                    else if (xlsFileName.startsWith("BOP180"))
                        writeXmlBOP180();
                    else if (xlsFileName.startsWith("PMI120"))
                        writeXmlPMI120();
                    else if (xlsFileName.startsWith("STA-STV"))
                        writeXmlSTASTV();
                    else if (xlsFileName.startsWith("FI601"))
                        writeXmlFI601();
                    else if (xlsFileName.startsWith("FI401"))
                        writeXmlFI401();
                    else if (xlsFileName.startsWith("ARM"))
                        writeXmlARM();
                    else if(xlsFileName.startsWith("STM231"))
                        writeXmlSTM231();
                    else if(xlsFileName.startsWith("MACHETA_RPS100"))
                        writeXmlRPS100();

                }
            }

            if(xlsFileId.equals("17") | (xlsFileId.equals("11") && xlsFileName.startsWith("MACHETA")) || (xlsFileId.equals("12") && xlsFileName.startsWith("SPD")) || (xlsFileId.equals("13") && xlsFileName.startsWith("SPE")) || (xlsFileId.equals("20") && xlsFileName.startsWith("SPD")) || (xlsFileId.equals("21") && xlsFileName.startsWith("SPD")) || xlsFileId.equals("22") || xlsFileId.equals("23") || xlsFileId.equals("25") || xlsFileId.equals("28") || xlsFileId.equals("29") || xlsFileName.startsWith("ASF610") || xlsFileName.startsWith("PVL110") || xlsFileName.startsWith("BOP180")|| xlsFileName.startsWith("MACHETA_RPS"))
            {
                writeXml("	</Body>");

                writeXml("</Message>");
            }
            else if (xlsFileId.equals("30"))
            {
                //writeXml("    </ReportedPayee>");
                writeXml("  </PaymentDataBody>");
                writeXml("</CESOP>");
            }
            else if (xlsFileId.equals("34"))
            {
                writeXml("    </Appendix>");
                writeXml("	</Body>");
                writeXml("</Message>");
            }
            else if (xlsFileId.equals("37"))
            {
                writeXml("		</IC>");
                writeXml("	</Body>");
                writeXml("</Message>");
            }
            else if (xlsFileId.equals("44"))
            {
                writeXml("</Participant-ReconImport>");
            }
            else if (xlsFileId.equals("41") || xlsFileId.equals("42")|| xlsFileId.equals("47"))
            {
                writeXml("	</Body>");
                writeXml("</Message>");
            }
            else if(xlsFileName.startsWith("ARM")) {
                writeXml("			</FinInstrmRptgTxRpt>\r\n"
                        + "		</Document>\r\n"
                        + "	</Pyld>\r\n"
                        + "</BizData>\r\n"
                        + "");
            }
            else if (!xlsFileId.equals("37") && !xlsFileId.equals("38") && !xlsFileId.equals("39") && !xlsFileId.equals("45") && !xlsFileId.equals("43") && !xlsFileId.equals("58"))
            {
                writeXml("</document>");


                if(xlsFileId.equals("16"))
                {
                    writeXmlIMV("</document>");
                    writeXmlRM("</document>");

                }
            }

            //eliberez resursele
            xmlWriter.close();
            if(xmlWriterIMV != null)
                xmlWriterIMV.close();
            if(xmlWriterRM != null)
                xmlWriterRM.close();
            workbook.close();
            statement.close();
            connection.close();


            if(xlsFileId.equals("16"))
            {
                zipFiles(xmlNamesList, path + xlsFileName);
            }

            System.out.println("DONE after " + (System.currentTimeMillis() - startTime) / 1000.0f + "s");

            validateXml(xsdSchema , path + xmlFileName);



        }
        catch(Exception e)
        {
            e.printStackTrace(new PrintWriter(sw));
            System.err.println("Exception at sheet: " + currentSheetName + "\n" + sw.toString());
            System.err.println("Exception: " + connection + " xlsFile_id= " + xlsFileId + " xlsFileName= "+ xlsFileName + "\n" + sw.toString());
            System.err.println("Exception at sqlQuerry: " + connectionString + "\n" + sw.toString());
            xmlWriter.close();
        }
    }
}
