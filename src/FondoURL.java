import java.awt.Graphics;
import java.awt.Image;
import java.io.IOException;
import java.net.URL; // Necesitas esta clase para la URL
import javax.imageio.ImageIO;
import javax.swing.JPanel;

public class FondoURL extends JPanel {
    
    private Image imagenDeFondo;
    private final String URL_IMAGEN = "https://images.stockcake.com/public/5/4/5/54511c70-242e-4046-9a97-b0f3e491daaa_medium/intense-gym-battle-stockcake.jpg";

    public FondoURL() {
        try {
            URL url = new URL(URL_IMAGEN);
            
            // 2. Usar ImageIO.read(URL) para cargar la imagen
            imagenDeFondo = ImageIO.read(url);
            
        } catch (IOException e) {
            System.err.println("Error al cargar la imagen desde la URL: " + URL_IMAGEN);
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