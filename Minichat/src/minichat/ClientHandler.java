package minichat;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.Socket;

public class ClientHandler implements Runnable {
	private Socket socket;
    private BufferedReader reader;
    private PrintWriter writer;
    private String username;

    public ClientHandler(Socket socket) {
        this.socket = socket;

        try {
            reader = new BufferedReader(new InputStreamReader(socket.getInputStream()));
            writer = new PrintWriter(socket.getOutputStream(), true);
        } catch (IOException e) {
            System.out.println("ClientHandler error: " + e.getMessage());
        }
    }

    @Override
    public void run() {
        try {
            writer.println("Enter your username:");
            username = reader.readLine();

            System.out.println(username + " joined the chat.");
            ChatServer.broadcast(username + " joined the chat.", this);

            String message;

            while ((message = reader.readLine()) != null) {
                if (message.equalsIgnoreCase("/quit")) {
                    break;
                }

                System.out.println(username + ": " + message);
                ChatServer.broadcast(username + ": " + message, this);
            }

        } catch (IOException e) {
            System.out.println("Connection error: " + e.getMessage());
        } finally {
            closeConnection();
        }
    }

    public void sendMessage(String message) {
        writer.println(message);
    }

    private void closeConnection() {
        try {
            ChatServer.removeClient(this);
            ChatServer.broadcast(username + " left the chat.", this);

            if (reader != null) {
                reader.close();
            }

            if (writer != null) {
                writer.close();
            }

            if (socket != null) {
                socket.close();
            }

        } catch (IOException e) {
            System.out.println("Close connection error: " + e.getMessage());
        }
    }
}
