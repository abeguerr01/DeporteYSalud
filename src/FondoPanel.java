import java.awt.Graphics;
import java.awt.Image;
import java.io.IOException;
import javax.imageio.ImageIO;
import javax.swing.JPanel;

public class FondoPanel extends JPanel {
    
    private Image imagenDeFondo;
    
        private final String RUTA_IMAGEN = "imagenes/fondoMenu.jpg";

    public FondoPanel() {
        try {
            imagenDeFondo = ImageIO.read(getClass().getResource(RUTA_IMAGEN));
        } catch (IOException e) {
            System.err.println("Error al cargar la imagen de fondo: " + RUTA_IMAGEN);
            e.printStackTrace();
        }
    }

    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g); 
        
        if (imagenDeFondo != null) {
            g.drawImage(imagenDeFondo, 0, 0, getWidth(), getHeight(), this);
        }
    }
}