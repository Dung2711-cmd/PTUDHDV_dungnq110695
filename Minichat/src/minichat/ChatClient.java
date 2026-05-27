package minichat;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.ConnectException;
import java.net.Socket;

public class ChatClient {
    private static final String DEFAULT_SERVER_HOST = "localhost";
    private static final int SERVER_PORT = 9999;

    public static void main(String[] args) {
        String serverHost = args.length > 0 ? args[0] : DEFAULT_SERVER_HOST;
        int serverPort = args.length > 1 ? parsePort(args[1]) : SERVER_PORT;

        try (
            Socket socket = new Socket(serverHost, serverPort);
            BufferedReader serverReader = new BufferedReader(new InputStreamReader(socket.getInputStream()));
            PrintWriter serverWriter = new PrintWriter(socket.getOutputStream(), true);
            BufferedReader keyboardReader = new BufferedReader(new InputStreamReader(System.in))
        ) {
            System.out.println("Connected to MiniChat Server.");

            Thread receiveThread = new Thread(() -> {
                try {
                    String serverMessage;

                    while ((serverMessage = serverReader.readLine()) != null) {
                        System.out.println(serverMessage);
                    }

                } catch (IOException e) {
                    System.out.println("Disconnected from server.");
                }
            });

            receiveThread.start();

            String userInput;

            while ((userInput = keyboardReader.readLine()) != null) {
                serverWriter.println(userInput);

                if (userInput.equalsIgnoreCase("/quit")) {
                    break;
                }
            }

        } catch (IOException e) {
            if (e instanceof ConnectException) {
                System.out.println("Cannot connect to MiniChat Server at " + serverHost + ":" + serverPort + ".");
                System.out.println("Please run ChatServer first, then run ChatClient again.");
            } else {
                System.out.println("Client error: " + e.getMessage());
            }
        }
    }

    private static int parsePort(String value) {
        try {
            return Integer.parseInt(value);
        } catch (NumberFormatException e) {
            System.out.println("Invalid port '" + value + "'. Using default port " + SERVER_PORT + ".");
            return SERVER_PORT;
        }
    }
}
