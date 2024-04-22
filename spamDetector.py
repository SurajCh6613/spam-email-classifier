import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
from win32com.client import Dispatch

def add_bg_from_url():
    st.markdown(
         f"""
         <style>
         .stApp {{
             background-image: url("https://www.dice-comms.co.uk/wp-content/uploads/fly-images/1487/Hero-Campaign-Management_-Benchmarking-Your-Email-Campaigns-scaled-2400x1200-c.jpg");
             background-attachment: fixed;
             background-size: cover
         }}
         </style>
         """,
         unsafe_allow_html=True
     )

add_bg_from_url()

def speak(text):
	speak = Dispatch("SAPI.SpVoice")
	speak.Speak(text)

model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

def main():
	st.title("Email Spam Classification Application")
	activites = ["Classification", "About"]
	choices = st.sidebar.selectbox("Select Activities", activites)
	if choices == "Classification":
		st.subheader("Classification")
		msg = st.text_area("Enter text here")
		if st.button("Predict"):
			print(msg)
			print(type(msg))
			data = [msg]
			print(data)
			vec = cv.transform(data).toarray()
			result = model.predict(vec)
			if result[0] == 0:
				st.success("This is Not A Spam Email")
				speak("This is Not A Spam Email")
			else:
				st.error("This is A Spam Email")
				speak("This is A Spam Email")
	if choices == "About":
		st.subheader("About")
		st.write("This is a spam email classifier application.")
		st.write("Made with streamlit")
		st.write("By Suraj Chaudhary and Vipin Singh")

main()
