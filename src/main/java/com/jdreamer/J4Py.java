package com.jdreamer;

import py4j.GatewayServer;

public class J4Py {
    public static void main(String[] args) {
        J4Py app = new J4Py();
        GatewayServer server = new GatewayServer(app);
        server.start();
    }
}
