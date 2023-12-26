import 'package:flutter/material.dart';
import 'web_view.dart';

class HomePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      extendBodyBehindAppBar: true,
      appBar: AppBar(
        backgroundColor: Colors.transparent,
        elevation: 0,
      ),
      body: OrientationBuilder(
        builder: (context, orientation) {
          return SingleChildScrollView(
            child: Container(
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
                    Image.asset( //Starting copy here
                      'images/murder_mafia.gif', // Replace with the actual image path
                      height: 200, // Set the desired height
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
                            title: "Murder Mafia",
                            selectedUrl: "https://www.y8.com/games/murder_mafia",
                          ),
                        ));
                      },
                      child: Text(
                        "Murder Mafia",
                        style: TextStyle(fontSize: 18.0),
                      ),
                    ),
                    SizedBox(height: 20), //End Copy Here
                      Image.asset(
                      'images/drunken_boxing.gif', // Replace with the actual image path
                      height: 200, // Set the desired height
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
                            title: "Drunken Boxing 2P",
                            selectedUrl: "https://www.y8.com/games/drunken_boxing",
                          ),
                        ));
                      },
                      child: Text(
                        "Drunken Boxing 2P",
                        style: TextStyle(fontSize: 18.0),
                      ),
                    ),
                    SizedBox(height: 20),
                    Image.asset(
                      'images/murder_mafia.gif', // Replace with the actual image path
                      height: 200, // Set the desired height
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
                        "2nd Game",
                        style: TextStyle(fontSize: 18.0),
                      ),
                    ),
                  ],
                ),
              ),
            ),
          );
        },
      ),
    );
  }
}
