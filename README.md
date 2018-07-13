# PowerPoint-OSC
A Powerpoint add-in for basic control of powerpoint presentations with OSC

## Installation

- Download and extract the add-in from here : http://benjamin.kuperberg.fr/download/powerpoint-osc.zip
- Launch **install.bat** (You may have to run it as administrator if you get an error message)
- Launch PowerPoint, you should see an "OSC" Tab, where you can configure the host and port to receive and send OSC messages.

## Usage
**You can only control the slides when the Slideshow is active, not in the editor !**

You can control slides by using :
- **/next** Next slide
- **/previous** Previous slide
- **/page <int>** Go to a specific slide (ex: /page 1)

Also, when the slide has changed, Powerpoint will send a /page <int> message as well with the current slide index.
