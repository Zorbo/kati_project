package run;
import process.XlsxBase;

import javax.swing.*;
import java.awt.*;
import java.io.IOException;

public class StartProcess {

    private static void createAndShowGUI() {

        JFrame.setDefaultLookAndFeelDecorated(true);
        JFrame frame = new JFrame("Bundle Example");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JPanel pane = new JPanel(new GridLayout(0, 1));
        JButton button = new JButton("Start!");
        pane.add(button);
        button.addActionListener(actionEvent -> {
            XlsxBase xlsxBase = new XlsxBase();
            try {
                xlsxBase.readXlsx("2019","October");
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                xlsxBase.writeXlsx();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
        JLabel label = new JLabel("Example for Bundling JRE with Java Class");
        pane.add(label);
        pane.setBorder(BorderFactory.createEmptyBorder(200, 200, 50, 200));
        frame.getContentPane().add(pane);
        frame.pack();
        frame.setVisible(true);
    }


    public static void main(String[] args) throws IOException {

        SwingUtilities.invokeLater(StartProcess::createAndShowGUI);
//        XlsxBase xlsxBase = new XlsxBase();
//        xlsxBase.readXlsx("2019","October");
//        xlsxBase.writeXlsx();
    }
}
