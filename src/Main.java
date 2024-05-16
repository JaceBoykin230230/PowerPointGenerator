import com.fasterxml.jackson.core.StreamReadFeature;
import com.fasterxml.jackson.databind.JsonDeserializer;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.xslf.usermodel.*;

import java.io.DataInput;
import java.io.FileOutputStream;
import java.util.*;
import java.util.random.RandomGenerator;

import java.io.File;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args) {

        //read from json
        ObjectMapper inputJson = new ObjectMapper();
        try {
            Map<?, ?> file = inputJson.readValue(new File("C:\\Users\\jaceb\\IdeaProjects\\CreatePowerPoint\\src\\pptPrompt.json"), Map.class);

            String title = file.get("title").toString();

            int slideCountInt = (int) file.get("slidecount");
            String author = file.get("author").toString();
            String date = file.get("date").toString();
            String startingcontent = file.get("startingcontent").toString();
            String audience = file.get("audience").toString();

            String prompt = "Make a powerpoint about " + title + " for this audience: " + audience + ", with " +
                    slideCountInt + " slides. The author is " + author + " and the date is " + date +
                    ". Store both the author and date in the JSON format. The starting content is " + startingcontent + ". The starting content should be considered the main theme of the presentation and each slide should be based around that. Remember, the content for the presentation should be tailored to the specific audience. Fill in details of the powerpoint for each slide. " +
                    "Format each slide with a brief title section and a content section that includes the relevant information. Do not number each slide in its title. Make the content information short enough to where it can fit in just one or two lines on a powerpoint. " +
                    "Make the whole response returned in JSON format. put the title and content section in a list called 'slides'." ;

            try {
                String gptResponse = OpenAI.chatGPT(prompt);

                //write to json
                ObjectMapper readJson = new ObjectMapper();
                readJson.enable(SerializationFeature.INDENT_OUTPUT);
                JsonNode jsonNode = readJson.readTree(gptResponse);
                readJson.writeValue(new File("C:\\Users\\jaceb\\IdeaProjects\\CreatePowerPoint\\src\\pptResponse.json"), jsonNode);

                //read from json
                ObjectMapper mapper = new ObjectMapper();
                mapper.enable(SerializationFeature.INDENT_OUTPUT);
                JsonNode file2 = mapper.readValue(new File("C:\\Users\\jaceb\\IdeaProjects\\CreatePowerPoint\\src\\pptResponse.json"), JsonNode.class);

                //get data from json to create powerpoint
                JsonNode content = mapper.readTree(file2.asText());
                String authorResponse = content.get("author").asText();
                System.out.println(authorResponse);
                String dateResponse = content.get("date").asText();
                System.out.println(dateResponse);

                JsonNode slidesResponse = content.get("slides");
                System.out.println(slidesResponse);

                //put all the data in a hashmap
                Map<String, String> slideMap = new LinkedHashMap<>();
                slideMap.put("author", authorResponse);
                slideMap.put("date", dateResponse);

                if(slidesResponse.isArray()){
                    int i = 1;
                    for(JsonNode slide : slidesResponse) {
                        if (slide != null) {
                            String titleResponse = slide.get("title").asText();
                            System.out.println(titleResponse);
                            slideMap.put("title"+ i, titleResponse);

                            String contentResponse = slide.get("content").asText();
                            System.out.println(contentResponse);
                            slideMap.put("content"+ i, contentResponse);
                            i++;
                        }
                    }

                    for (Map.Entry<String, String> entry : slideMap.entrySet()) {
                        System.out.println(entry.getKey() + " = " + entry.getValue());
                    }
                }

                //create powerpoint
                XMLSlideShow ppt = new XMLSlideShow();
                //default slide
                XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);

                // title slide layout
                XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
                XSLFTheme theme2 = titleLayout.getTheme();
                titleLayout.getTheme();
                System.out.println(theme2.getName());


                //slide layout
                XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
                XSLFTheme theme = layout.getTheme();
                layout.getTheme();
                System.out.println(theme.getName());


                XSLFSlide nextSlide = ppt.createSlide(titleLayout);
                XSLFTextShape titleShape = nextSlide.getPlaceholder(0);
                titleShape.clearText();
                titleShape.setText(slideMap.get("title" + (1)));
                XSLFTextBox authorBox = nextSlide.createTextBox();



                XSLFTextShape sampleContent = nextSlide.getPlaceholder(1);
                sampleContent.clearText();
                sampleContent.addNewTextParagraph().addNewTextRun().setText(authorResponse);
                sampleContent.addNewTextParagraph().addNewTextRun().setText(dateResponse);



                for (int i = 1; i <= slideCountInt; i++) {

                     if (i > 1) {
                        // create the rest of slides
                        XSLFSlide nextSlide2 = ppt.createSlide(layout);
                        XSLFTextShape titleShape2 = nextSlide2.getPlaceholder(0);
                        titleShape2.clearText();
                        titleShape2.setText(slideMap.get("title" + (i)));
                        XSLFTextShape contentShape = nextSlide2.getPlaceholder(1);
                        contentShape.clearText();

                        contentShape.setText(slideMap.get("content" + (i)));


                        //This code adds more bullet points to a powerpoint slide 20% of the time, for more detail.
                        double rand = Math.random();
                        if (rand < 0.3) {
                            //add author
                            String prompt2 = "Take this content: " + slideMap.get("content" + (i)) + " and title: " + slideMap.get("title" + (i)) + " and elaborate on the content in a subcontent section called 'subcontent'. keep each point in the subcontent brief because it will be used inside a powerpoint." +
                                    " return the subcontent in json format, with 'subcontent'  being the key. the value, being each point inside of a list . " +
                                    "Do not add more than two points under the subcontent section, there should only be one or two points at most so it can all fit in a powerpoint slide. Make the points brief. Try not to repeat details that have already been used in the content or other subcontent sections.";
                            String gptResponse2 = OpenAI.chatGPT(prompt2);
                            ObjectMapper readJson2 = new ObjectMapper();
                            readJson2.enable(SerializationFeature.INDENT_OUTPUT);
                            JsonNode jsonNode2 = readJson2.readTree(gptResponse2);
                            readJson2.writeValue(new File("C:\\Users\\jaceb\\IdeaProjects\\CreatePowerPoint\\src\\pptResponse2.json"), jsonNode2);

                            ObjectMapper mapper2 = new ObjectMapper();
                            mapper2.enable(SerializationFeature.INDENT_OUTPUT);
                            JsonNode file3 = mapper2.readValue(new File("C:\\Users\\jaceb\\IdeaProjects\\CreatePowerPoint\\src\\pptResponse2.json"), JsonNode.class);
                            JsonNode content2 = mapper2.readTree(file3.asText());
                            JsonNode subcontent = content2.get("subcontent");
                            //System.out.println(authorResponse2);
                            Map<String, String> subcontentMap = new HashMap<>();
                            int x = 0;
                            for(JsonNode subcontents: subcontent) {
                                if (subcontents != null) {
                                    String subcontentResponse = subcontent.get(x).asText();
                                    System.out.println(subcontentResponse);
                                    subcontentMap.put("subcontent" + x, subcontentResponse);
                                    x++;
                                }
                            }
                            for (Map.Entry<String, String> entry : subcontentMap.entrySet()) {
                                contentShape.addNewTextParagraph().addNewTextRun().setText(entry.getValue());
                            }
                        }
                    }
                }

                //Save the Presentation
                FileOutputStream outputStream = new FileOutputStream("C:\\Users\\jaceb\\IdeaProjects\\CreatePowerPoint\\src\\samplePPT.pptx");
                ppt.write(outputStream);
                outputStream.close();

                // Closing presentation
                ppt.close();

            } catch (Exception e) {
                System.out.println(e);
            }

        }catch (Exception e){
            System.out.println(e);
        }
    }
}