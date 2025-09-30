from kivy.app import App
from kivy.uix.button import Button

class MyApp(App):
    def build(self):
        # Create a button widget
        return Button(
            text="Hello, Kivy!",
            font_size=32,
            on_press=self.on_button_click
        )

    def on_button_click(self, instance):
        # Change the button text when clicked
        instance.text = "You clicked me!"

if __name__ == "__main__":
    MyApp().run()
