package xLsToXmL;

import javax.swing.SwingUtilities;

public class MainClass {

    public static void main(String[] args) {
        InterfaceClass ui = new InterfaceClass();
        SwingUtilities.invokeLater(ui);
    }
}