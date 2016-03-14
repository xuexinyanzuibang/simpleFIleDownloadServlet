package cn.com.genechem.doc;

import java.net.URL;

public class DocGenerator {
    URL template;

    public DocGenerator(URL template) {
        this.template = template;

    }
    public void generate(DocModel model, String output) throws Exception {
        new DocResult(template,model).save(output);
    }
}
