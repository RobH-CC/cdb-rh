cardcontent = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.2",
    "body": [
        {
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "width": 2,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Sign up for our Newsletter!",
                            "weight": "Bolder",
                            "size": "Medium"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Our newsletter is super freaking awesome, and there is no other like it in the universe.",
                            "isSubtle": True,
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Don't worry, we'll never share or sell your information (unless we can profit from it hehe).",
                            "isSubtle":True,
                            "wrap": True,
                            "size": "Small"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Your name",
                            "wrap": True
                        },
                        {
                            "type": "Input.Text",
                            "id": "myName",
                            "placeholder": "First and last name"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Your email",
                            "wrap": True
                        },
                        {
                            "type": "Input.Text",
                            "id": "myEmail",
                            "placeholder": "youremail@example.com",
                            "style": "Email"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Phone Number"
                        },
                        {
                            "type": "Input.Text",
                            "id": "myTel",
                            "placeholder": "000-000-0000",
                            "style": "Tel"
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "Image",
                            "url": "https://i.postimg.cc/wMJvqNR6/sign-up.jpg",
                            "size": "auto"
                        }
                    ]
                }
            ]
        }
    ],
    "actions": [
        {
            "type": "Action.Submit",
            "title": "Submit"
        }
    ]
}
