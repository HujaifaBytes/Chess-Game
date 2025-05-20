import tkinter as tk
from tkinter import font
import random
import time

class RockPaperScissorsGame:
    def __init__(self, root):
        self.root = root
        self.root.title("Rock Paper Scissors")
        self.root.geometry("900x800") # Increased height slightly for more history
        self.root.configure(bg="#f5f0ff") 

        # Game state
        self.player_score = 0
        self.computer_score = 0
        self.game_history = [] # Will now store full history
        self.is_animating = False
        self.max_history_display = 7 # Number of history items to show on screen

        # Choices and Emojis
        self.choices = ['rock', 'paper', 'scissors']
        self.emojis = {
            'rock': '👊',
            'paper': '✋',
            'scissors': '✌️',
            'thinking': '🤔', 
            'question': '❓'  
        }
        self.winning_rules = {
            'rock': 'scissors',
            'paper': 'rock',
            'scissors': 'paper'
        }

        # Fonts
        self.title_font = font.Font(family="Arial", size=28, weight="bold")
        self.score_label_font = font.Font(family="Arial", size=14, weight="bold")
        self.score_font = font.Font(family="Arial", size=32, weight="bold")
        self.player_label_font = font.Font(family="Arial", size=16, weight="bold")
        self.hand_font = font.Font(family="Arial", size=50) 
        self.result_font = font.Font(family="Arial", size=24, weight="bold")
        self.button_font = font.Font(family="Arial", size=30) 
        self.history_title_font = font.Font(family="Arial", size=16, weight="bold")
        self.history_item_font = font.Font(family="Arial", size=12)

        # UI Colors
        self.primary_color = "#6a26cd" 
        self.secondary_color = "#8a4bdf" 
        self.text_color_dark = "#444"
        self.win_color = "#2ecc71" 
        self.lose_color = "#e74c3c" 
        self.tie_color = "#3498db" 
        self.container_bg = "#f0e6ff" 
        self.card_bg = "white"

        self.create_widgets()

    def create_widgets(self):
        # Main game container
        self.game_container = tk.Frame(self.root, bg=self.container_bg, padx=20, pady=20)
        self.game_container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        # Title
        title_label = tk.Label(self.game_container, text="Rock Paper Scissors", font=self.title_font, fg=self.primary_color, bg=self.container_bg)
        title_label.pack(pady=(0, 20))

        # Score Board
        score_board_frame = tk.Frame(self.game_container, bg=self.container_bg)
        score_board_frame.pack(fill=tk.X, pady=(0, 20))

        # Player Score Card
        player_score_card = tk.Frame(score_board_frame, bg=self.card_bg, relief=tk.SOLID, bd=1, borderwidth=0, highlightbackground="#e0e0e0", highlightthickness=1)
        player_score_card.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,10))
        tk.Label(player_score_card, text="You", font=self.score_label_font, fg=self.primary_color, bg=self.card_bg).pack(pady=(10,0))
        self.player_score_label = tk.Label(player_score_card, text="0", font=self.score_font, fg=self.text_color_dark, bg=self.card_bg)
        self.player_score_label.pack(pady=(0,10))

        # Computer Score Card
        computer_score_card = tk.Frame(score_board_frame, bg=self.card_bg, relief=tk.SOLID, bd=1, borderwidth=0, highlightbackground="#e0e0e0", highlightthickness=1)
        computer_score_card.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(10,0))
        tk.Label(computer_score_card, text="Computer", font=self.score_label_font, fg=self.primary_color, bg=self.card_bg).pack(pady=(10,0))
        self.computer_score_label = tk.Label(computer_score_card, text="0", font=self.score_font, fg=self.text_color_dark, bg=self.card_bg)
        self.computer_score_label.pack(pady=(0,10))

        # Game Area (Hands)
        game_area_frame = tk.Frame(self.game_container, bg=self.container_bg)
        game_area_frame.pack(fill=tk.X, pady=20)

        # Player Hand
        player_box = tk.Frame(game_area_frame, bg=self.container_bg)
        player_box.pack(side=tk.LEFT, expand=True)
        tk.Label(player_box, text="You", font=self.player_label_font, fg=self.text_color_dark, bg=self.container_bg).pack()
        self.player_hand_label = tk.Label(player_box, text=self.emojis['question'], font=self.hand_font, bg=self.card_bg, width=3, height=1, relief=tk.SOLID, bd=1, borderwidth=0, highlightbackground="#e0e0e0", highlightthickness=1)
        self.player_hand_label.pack(pady=10)

        # VS Label
        vs_label = tk.Label(game_area_frame, text="VS", font=self.result_font, fg=self.primary_color, bg=self.container_bg)
        vs_label.pack(side=tk.LEFT, expand=True, padx=20)

        # Computer Hand
        computer_box = tk.Frame(game_area_frame, bg=self.container_bg)
        computer_box.pack(side=tk.RIGHT, expand=True)
        tk.Label(computer_box, text="Computer", font=self.player_label_font, fg=self.text_color_dark, bg=self.container_bg).pack()
        self.computer_hand_label = tk.Label(computer_box, text=self.emojis['question'], font=self.hand_font, bg=self.card_bg, width=3, height=1, relief=tk.SOLID, bd=1, borderwidth=0, highlightbackground="#e0e0e0", highlightthickness=1)
        self.computer_hand_label.pack(pady=10)

        # Result Message
        self.result_label = tk.Label(self.game_container, text="", font=self.result_font, bg=self.container_bg, height=2)
        self.result_label.pack(pady=10)

        # Controls (Buttons)
        controls_frame = tk.Frame(self.game_container, bg=self.container_bg)
        controls_frame.pack(pady=20)

        self.rock_btn = tk.Button(controls_frame, text=self.emojis['rock'], font=self.button_font, bg=self.secondary_color, fg="white", relief=tk.RAISED, bd=2, width=3, command=lambda: self.play_round('rock'), cursor="hand2")
        self.rock_btn.pack(side=tk.LEFT, padx=10)
        self.paper_btn = tk.Button(controls_frame, text=self.emojis['paper'], font=self.button_font, bg=self.secondary_color, fg="white", relief=tk.RAISED, bd=2, width=3, command=lambda: self.play_round('paper'), cursor="hand2")
        self.paper_btn.pack(side=tk.LEFT, padx=10)
        self.scissors_btn = tk.Button(controls_frame, text=self.emojis['scissors'], font=self.button_font, bg=self.secondary_color, fg="white", relief=tk.RAISED, bd=2, width=3, command=lambda: self.play_round('scissors'), cursor="hand2")
        self.scissors_btn.pack(side=tk.LEFT, padx=10)

        # History Container
        history_outer_frame = tk.Frame(self.game_container, bg=self.card_bg, relief=tk.SOLID, bd=1, borderwidth=0, highlightbackground="#e0e0e0", highlightthickness=1)
        # Allow history to expand a bit more if needed, but not excessively
        history_outer_frame.pack(fill=tk.X, pady=(10,0), expand=False) 
        tk.Label(history_outer_frame, text="Game History (Last " + str(self.max_history_display) + ")", font=self.history_title_font, fg=self.primary_color, bg=self.card_bg).pack(pady=5)
        self.history_frame = tk.Frame(history_outer_frame, bg=self.card_bg)
        self.history_frame.pack(fill=tk.X, padx=10, pady=(0,10))
        self.update_history_display() 

    def set_button_state(self, state):
        """Enable or disable choice buttons."""
        self.rock_btn.config(state=state)
        self.paper_btn.config(state=state)
        self.scissors_btn.config(state=state)

    def play_round(self, player_choice):
        if self.is_animating:
            return
        self.is_animating = True
        self.set_button_state(tk.DISABLED)

        self.player_hand_label.config(text=self.emojis['question'], fg=self.text_color_dark)
        self.computer_hand_label.config(text=self.emojis['question'], fg=self.text_color_dark)
        self.result_label.config(text="", fg=self.text_color_dark)

        # Start countdown animation
        self.countdown_animation(3, player_choice)

    def countdown_animation(self, count, player_choice):
        if count > 0:
            self.player_hand_label.config(text=str(count), font=self.hand_font) 
            self.computer_hand_label.config(text=str(count), font=self.hand_font)
            self.root.after(600, lambda: self.countdown_animation(count - 1, player_choice))
        else:
            # After countdown, show player's choice and start computer "thinking"
            self.player_hand_label.config(text=self.emojis[player_choice], font=self.hand_font)
            self.computer_thinking_animation(player_choice)

    def computer_thinking_animation(self, player_choice, iterations=5):
        """Simple animation for computer's choice."""
        if iterations > 0:
            temp_emoji = random.choice([self.emojis['rock'], self.emojis['paper'], self.emojis['scissors'], self.emojis['thinking']])
            self.computer_hand_label.config(text=temp_emoji, font=self.hand_font)
            self.root.after(150, lambda: self.computer_thinking_animation(player_choice, iterations - 1))
        else:
            self.determine_winner(player_choice)


    def determine_winner(self, player_choice):
        computer_choice = random.choice(self.choices)

        self.player_hand_label.config(text=self.emojis[player_choice])
        self.computer_hand_label.config(text=self.emojis[computer_choice])

        result_text = ""
        result_color = self.text_color_dark

        if player_choice == computer_choice:
            result_text = "It's a Tie!"
            result_color = self.tie_color
        elif self.winning_rules[player_choice] == computer_choice:
            result_text = "You Win!"
            result_color = self.win_color
            self.player_score += 1
        else:
            result_text = "Computer Wins!"
            result_color = self.lose_color
            self.computer_score += 1
        
        self.result_label.config(text=result_text, fg=result_color)
        self.player_score_label.config(text=str(self.player_score))
        self.computer_score_label.config(text=str(self.computer_score))

        self.add_to_history(player_choice, computer_choice, result_text)
        
        self.root.after(1000, self.reset_for_next_round) 

    def reset_for_next_round(self):
        self.is_animating = False
        self.set_button_state(tk.NORMAL)

    def add_to_history(self, p_choice, c_choice, result):
        # Add to the beginning of the list to show most recent first
        self.game_history.insert(0, {"player": p_choice, "computer": c_choice, "result": result})
        # No longer popping items, self.game_history stores all rounds.
        # The display will be limited by update_history_display.
        self.update_history_display()

    def update_history_display(self):
        # Clear previous history items
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        if not self.game_history:
             tk.Label(self.history_frame, text="No games played yet.", font=self.history_item_font, bg=self.card_bg, fg=self.text_color_dark).pack(pady=5)
             return

        # Display only the most recent 'max_history_display' items
        # Slicing: game_history is already newest first due to insert(0, ...)
        history_to_show = self.game_history[:self.max_history_display] 

        for item in history_to_show:
            history_item_frame = tk.Frame(self.history_frame, bg=self.card_bg)
            history_item_frame.pack(fill=tk.X, pady=2)

            p_emoji_label = tk.Label(history_item_frame, text=self.emojis[item['player']], font=self.history_item_font, bg=self.card_bg)
            p_emoji_label.pack(side=tk.LEFT, padx=5)
            
            result_text = item['result'].replace("!", "") 
            res_color = self.text_color_dark
            if "You Win" in item['result']: res_color = self.win_color
            elif "Computer Wins" in item['result']: res_color = self.lose_color
            elif "Tie" in item['result']: res_color = self.tie_color

            result_label_hist = tk.Label(history_item_frame, text=f"- {result_text} -", font=self.history_item_font, bg=self.card_bg, fg=res_color)
            result_label_hist.pack(side=tk.LEFT, expand=True) # Allow this to take available space

            c_emoji_label = tk.Label(history_item_frame, text=self.emojis[item['computer']], font=self.history_item_font, bg=self.card_bg)
            c_emoji_label.pack(side=tk.RIGHT, padx=5)

if __name__ == "__main__":
    root = tk.Tk()
    game = RockPaperScissorsGame(root)
    root.mainloop()
