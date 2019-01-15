package run;
import process.XlsxBase;

import javax.swing.*;
import java.awt.*;
import java.io.IOException;

public class StartProcess {

    private static void createAndShowGUI() {

        // Create the panel
        JFrame.setDefaultLookAndFeelDecorated(true);
        JFrame frame = new JFrame("Ho vegi osszesito");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JPanel pane = new JPanel(new GridLayout(1, 3));
        JButton button = new JButton("Start!");


        // Set input fields

        JLabel yearLabel = new JLabel("Év: ", JLabel.TRAILING);
        JLabel monthLabel = new JLabel("Hónap: ", JLabel.TRAILING);

        JTextField year = new JTextField(10 );
        JTextField month = new JTextField(10);

        yearLabel.setLabelFor(year);
        monthLabel.setLabelFor(month);
        pane.add(yearLabel);
        pane.add(year);
        pane.add(monthLabel);
        pane.add(month);
        pane.add(button);

        // Add the button action listener
        button.addActionListener(actionEvent -> {
            XlsxBase xlsxBase = new XlsxBase();
            try {
                if (year.getText().isEmpty() || month.getText().isEmpty()) {
                    JOptionPane.showMessageDialog(frame,"Kérem adja meg az évszámot és a hónapot!");
                    return;
                } else {
                    xlsxBase.readXlsx(year.getText(),month.getText());
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                xlsxBase.writeXlsx();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
        JLabel label = new JLabel("Havi összesítés elkészítése");
        pane.add(label);
        pane.setBorder(BorderFactory.createEmptyBorder(200, 200, 50, 200));
        frame.getContentPane().add(pane);
        frame.pack();
        frame.setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(StartProcess::createAndShowGUI);
    }
}
