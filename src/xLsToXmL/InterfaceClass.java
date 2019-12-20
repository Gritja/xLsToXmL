package xLsToXmL;

import java.awt.Container;
import java.awt.Dimension;
import java.awt.GridLayout;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JTextArea;
import javax.swing.WindowConstants;

public class InterfaceClass implements Runnable {

	private JFrame frame;

	public InterfaceClass() {

	}

	@Override
	public void run() {
		frame = new JFrame("xLsToXmL");
		frame.setPreferredSize(new Dimension(600, 100));

		frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		createComponents(frame.getContentPane());

		frame.pack();
		frame.setVisible(true);
	}

	private void createComponents(Container container) {

		GridLayout layout = new GridLayout(2, 1);
		container.setLayout(layout);
		JTextArea textAreaBottom = new JTextArea();
		JButton converterButton = new JButton("Pick a file (.xls or .xlsx) to convert into XML!");
		ConvertListener converterAction = new ConvertListener(textAreaBottom);
		converterButton.addActionListener(converterAction);

		container.add(converterButton);
		container.add(textAreaBottom);

	}

	public JFrame getFrame() {
		return frame;
	}

}