---
toc: false
layout: post
description: Presentation are usually boring. What if you could change that. This is a blog post of how you can create custom data visualization in PowerPoint Using Python.
categories: [Python, PowerPoint, Data Visualization, OpenXML]
title: Building Data Visualizations in PowerPoint Using Python
sticky_rank: 1
image_banner: https://images.pexels.com/photos/261662/pexels-photo-261662.jpeg?auto=compress&cs=tinysrgb&dpr=2&h=650&w=940
---

The role of data to make decisions in organizations across all the sectors has changed the structural design of our organizations. Our need to be accurate in our decision making has emphasized the need to understand data better. To fulfill the need, the technology around Data Visualizations in the past couple of years has evolved by leaps and bounds.

Thus allowing developers and designers to build great visual artifacts helping people to expand the outreach of the messages they want to share. But the progress in the area of Data Visualization hasn’t been the same across all the platforms. One of these platforms which have missed the fruits of the progress in Powerpoint Presentation.

## Why Microsoft Powerpoint is an important platform?

Presentation is still one of the most popular medium that people use to communicate insights. Microsoft Powerpoint as a platform still holds a huge market base that spans across a diverse set of sectors. While the Powerpoint UI has evolved a lot since it first appeared in 1987, there have been powerpoint features that didn’t get to see equitable progress. Two such major areas, I believe are – Animations and Charts.

## Data Visualizations in Powerpoint

Visualization gives you answers to questions you didn’t know you had – Ben Schneiderman. Visualization in Powerpoint has been majorly static over the year and the palate to choose from has also been fairly traditional and limiting. While a simple bar chart can always be enough to convey a message, over usage of these graphs and often not in the right context have reduced the excitement data visualizations can generate and thus often fail to captivate our audience’s attention.

It has become harder to capture the attention of our audience to convey the message we want to share. Even from a creator’s perspective, creating a Powerpoint Presentation hasn’t been an exciting platform thus, we often hear people attributing the activity as a boring endeavor. It wouldn’t be too harsh to comment that the world’s largest medium to present insights is probably also one of the hated platforms to engage in.

Hate might be a strong word, but the catchphrase we often hear near the water cooler – “Another Presentation?” might agree with it. While I am not sure if it would be apt to blame the platform for people making bad presentations, I believe it would be a good subject for debate on.

## Coming together of two different worlds – Python and Powerpoint

I believe a lot of the designs cannot be realized because of the limitation of the feature PowerPoint provides. For Example – Imaging a data visualization with 200 data points in a single slide where each data point is represented using an image first represented in bubble clusters.

Now imagine those data points being rearranged in bar chart fashion in the next slide. And while in a slideshow, these images(data point) move from their initial position to the next position in what is termed as a [Sandance Visualization](https://sanddance.js.org/) . Manually putting these data points is not humanly feasible. Such kinds of visualizations sadly aren’t imagined for Powerpoint. But what if I can tell you that dream isn’t too far fetched.

## Diving in –

Building a presentation using Python opens a whole new possibility of designs as with code you can leverage the power of Automation. To do this, we will be using a python library known as Python PPTX. While python-pptx is popularly used for its power to update and create a presentation in a usual workflow process, a lot of interesting use-cases can be thought of.

For this blog post, we will be making a [Racing Bar chart](https://www.youtube.com/watch?v=8NQvismUv9A) to visualize how the Population of different countries has changed over the years.

## A little background

Python-PPTX can work with any OpenXML based Presentation platform. Microsoft Powerpoint is formatted in XML to represent the content generated with an aim to ease the transfer of data across their platform. Please refer to the installation process of the Python PPTX package before you begin with code - [https://python-pptx.readthedocs.io/en/latest/user/install.html](https://python-pptx.readthedocs.io/en/latest/user/install.html).

## Lets Code –

#### Import Package

```python
# Import Objects from Python PPTX
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt, Cm, Inches
from pptx.enum.text import PP_ALIGN

# Miscellaneous Imports
#Pandas to read and process data
import pandas as pd
import math
```

#### Create an empty Presentation Object

https://python-pptx.readthedocs.io/en/latest/user/quickstart.html#hello-world-example

```python
prs = Presentation()
```

#### Choose a slide Layout type.

https://python-pptx.readthedocs.io/en/latest/dev/analysis/sld-layout.html

```python
title_slide_layout = prs.slide_layouts[0]
blank_slide_layout = prs.slide_layouts[6]
```

#### Add a Slide

https://python-pptx.readthedocs.io/en/latest/api/slides.html

```python
slide = prs.slides.add_slide(title_slide_layout)
```

#### Access Slide default objects

Set the title and subtitle for the visualization.

```python
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "World Population Over the Years"
subtitle.text = "A look into how human settlement has evolved."
```

### Read data from a dataset

```python
df = pd.read_csv("data.csv")
```

#### Get all the unique year for available in the dataset.

```python
years = df["Year"].unique()
```

#### Define a scale

Function to define a scale which provides us the width size in conjuction with the available width size.

```python
def scale(value, min_val, max_val, minScale, maxScale):
    return minScale + (value - min_val) / (max_val - min_val) \* (maxScale - minScale)
```

#### Define function to make numbers readable

Give that Population data goes from thousands to a couple billions. It is important to make the numbers readable to the audience.

```python
millnames = ['',' K',' M',' B',' T']

def millify(n):
    n = float(n)
    millidx = max(0,min(len(millnames)-1,
    int(math.floor(0 if n == 0 else math.log10(abs(n))/3))))
    return '{:.0f}{}'.format(n / 10**(3 * millidx), millnames[millidx])
```

#### Create Data Visualization objects

Each slide is a snapshot representation of a year. We create a scale function to get what would be an appropriate width of each. We use a couple of textbox showcase the country name and values.

```python

for year in years:

    # Add a slide for each year.
    slide = prs.slides.add_slide(blank_slide_layout)

    # Slice the main dataframe(dataset) by year and sort the sliced dataframe by the value
    df_slice_by_year = df[df["Year"] == year].sort_values("value", ascending =False)

    # Find the Maximum value in the data slice which would be the top row.
    slice_max = df_slice_by_year.iloc[0]["value"];

    # Set Min as 0. Bar charts should start from 0 unless you have negative values.
    slice_min = 0;

    #Get the SLide width
    slideWidth = prs.slide_width;

    # Add a text box to showcase represent each slide by year.
    text_box = slide.shapes.add_textbox(Cm(19), Cm(15), Cm(4), Cm(2))
    text_box.text = str(year)
    text_box.text_frame.paragraphs[0].font.size = Pt(54); # Change font size of the textbox
    text_box.text_frame.paragraphs[0].font.bold = True; # Change font weight of the textbox

    # loop to create bars in a single slide. For the this visualization we will only cover Top 15 countries for each year.
    # Here each row in the dataset represents Country and their corresponding Population.
    for index, country in enumerate(df_slice_by_year.head(15).iterrows()):

        # Value which becomes the basis for measurement
        value = country[1]["value"]

        # Setup the Maximum width based on the current slide width
        scaleMaxVal = slideWidth * 0.7;

        # Get the width of each bar which has been scaled to the width of the slide.
        scaledValue = scale(value, 0, slice_max, 0, scaleMaxVal);

        # Add the Bar textbox. Please refer the function documentation to understand what params are passed.
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
            Cm(5.5), Cm(( 0.7 * (index + 1) + (0.21 * (index + 1))) + 2), scaledValue  , Cm(0.7));

        # We add a reference name to each object for us to create a nice motion graphic.
        bar.name = "!!" + str(country[1]["Country"]).replace(" ", "").replace("[^0-9a-zA-Z]+", "")

        # Add the value towards the end of the bar to help audience get the order of magnitude.
        text_frame = bar.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT # Choose Alignment
        p.font.size =  Pt(10) # Change font size
        p.text = millify(country[1]["value"]) # Format the number to human readable form

        # Add the country name on the Y-Axis to help user track the country.
        bartext = slide.shapes.add_textbox(Cm(2), Cm(( 0.7 * (index + 1) + (0.21 * (index + 1))) + 2), Cm(1.5), Cm(1.5));
        bartext_run = bartext.text_frame.add_paragraph()
        bartext.text = country[1]["Country"]
        bartext.text_frame.paragraphs[0].font.size = Pt(12);
        bartext_run.font.bold = True;

# Save the presentation.
prs.save('Raching Bar Chart.pptx')

```

#### After successfully generating the presentation, follow the following steps

- Open the generated presentation.
- Select all the slides.
- Choose "Morph" animation in the Transition Tab (Only Available in Office 2016 version and above).
- Change the duration of each slide to 0.5 (Feel free to set it up based on your convenience).
- In the Advance Slide section, deselect the "On Mouse Click" option.
- In the same section select the "After" option.
- Voila. Open the presentation in SlideShow mode.
