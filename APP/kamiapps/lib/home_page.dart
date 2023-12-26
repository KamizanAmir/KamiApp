import 'package:flutter/material.dart';
import 'web_view.dart';

class HomePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      extendBodyBehindAppBar: true, // Allow the body to be behind the app bar
      appBar: AppBar(
        backgroundColor: Colors.transparent, // Set app bar background color to transparent
        elevation: 0, // Remove the shadow under the app bar
      ),
      body: OrientationBuilder(
        builder: (context, orientation) {
          return Container(
            decoration: BoxDecoration(
              image: DecorationImage(
                image: AssetImage('images/aic_icon.jpg'),
                fit: BoxFit.cover,
              ),
            ),
            child: Center(
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  TextButton(
                    style: TextButton.styleFrom(
                      primary: Colors.white,
                      backgroundColor: Colors.blue.withOpacity(0.5),
                      padding: EdgeInsets.all(16.0),
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(8.0),
                      ),
                      minimumSize: Size(100, 0),
                    ),
                    onPressed: () {
                      Navigator.of(context).push(MaterialPageRoute(
                        builder: (BuildContext context) => MyWebView(
                          title: "MES System",
                          selectedUrl: "http://191.1.1.26/",
                        ),
                      ));
                    },
                    child: Text(
                      "Login Page",
                      style: TextStyle(fontSize: 18.0),
                    ),
                  ),
                  SizedBox(height: 20),
                  TextButton(
                    style: TextButton.styleFrom(
                      primary: Colors.white,
                      backgroundColor: Colors.blue.withOpacity(0.5),
                      padding: EdgeInsets.all(16.0),
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(8.0),
                      ),
                      minimumSize: Size(100, 0),
                    ),
                    onPressed: () {
                      Navigator.of(context).push(MaterialPageRoute(
                        builder: (BuildContext context) => MyWebView(
                          title: "AIC Website",
                          selectedUrl: "https://www.aicsemicon.com/",
                        ),
                      ));
                    },
                    child: Text(
                      "AIC",
                      style: TextStyle(fontSize: 18.0),
                    ),
                  ),
                ],
              ),
            ),
          );
        },
      ),
    );
  }
}
