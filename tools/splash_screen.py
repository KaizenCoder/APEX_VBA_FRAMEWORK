import tkinter as tk
from tkinter import ttk
import time
from pathlib import Path
from typing import Callable
import math

class AnimatedLabel(ttk.Label):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._text = kwargs.get('text', '')
        self._current_text = ''
        self._index = 0
        self._dots = 0
        self._animate_dots = False

    def start_animation(self):
        """Démarre l'animation du texte."""
        self._animate_text()

    def _animate_text(self):
        """Animation du texte lettre par lettre."""
        if self._index < len(self._text):
            self._current_text += self._text[self._index]
            self.config(text=self._current_text)
            self._index += 1
            self.after(50, self._animate_text)
        else:
            self._animate_dots = True
            self._animate_loading_dots()

    def _animate_loading_dots(self):
        """Animation des points de chargement."""
        if self._animate_dots:
            dots = '.' * self._dots
            self.config(text=f"{self._text}{dots}")
            self._dots = (self._dots + 1) % 4
            self.after(500, self._animate_loading_dots)

class AnimatedProgressbar(ttk.Progressbar):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._progress = 0
        self._target = 0
        self._speed = 2

    def set_target(self, target: float):
        """Définit la valeur cible avec animation."""
        self._target = target
        if not self._progress:
            self._animate_progress()

    def _animate_progress(self):
        """Animation fluide de la barre de progression."""
        if self._progress < self._target:
            diff = (self._target - self._progress) / self._speed
            self._progress += diff
            self['value'] = self._progress
            self.after(16, self._animate_progress)

class SplashScreen:
    def __init__(self, duration: int = 3):
        self.duration = duration
        self.window = tk.Tk()
        self.window.overrideredirect(True)
        self.window.attributes('-alpha', 0.0)
        
        # Centrer la fenêtre
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        width = 400
        height = 300
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        
        # Style
        self.window.configure(bg='#2B2D42')
        style = ttk.Style()
        style.configure('Splash.TLabel', 
                       background='#2B2D42', 
                       foreground='#EDF2F4',
                       font=('Segoe UI', 14))
        style.configure('Splash.TProgressbar', 
                       troughcolor='#8D99AE',
                       background='#EF233C')
        
        # Logo APEX avec animation de rotation
        self.canvas = tk.Canvas(self.window, 
                              width=100, 
                              height=100, 
                              bg='#2B2D42',
                              highlightthickness=0)
        self.canvas.pack(pady=20)
        self._create_animated_logo()
        
        # Titre animé
        self.title_label = AnimatedLabel(self.window, 
                                       text="APEX Orchestrator",
                                       style='Splash.TLabel')
        self.title_label.pack(pady=10)
        
        # Message de chargement animé
        self.status_label = AnimatedLabel(self.window, 
                                        text="Initialisation",
                                        style='Splash.TLabel')
        self.status_label.pack(pady=20)
        
        # Barre de progression animée
        self.progress = AnimatedProgressbar(self.window, 
                                          style='Splash.TProgressbar',
                                          length=300, 
                                          mode='determinate')
        self.progress.pack(pady=20)

    def _create_animated_logo(self):
        """Crée un logo APEX animé."""
        self.logo_angle = 0
        size = 40
        cx, cy = 50, 50
        
        # Points du A
        self.a_points = [
            cx-size/2, cy+size/2,  # base gauche
            cx, cy-size/2,         # sommet
            cx+size/2, cy+size/2   # base droite
        ]
        
        # Barre du A
        self.a_bar = [
            cx-size/4, cy,
            cx+size/4, cy
        ]
        
        self._draw_logo()
        self._animate_logo()

    def _draw_logo(self):
        """Dessine le logo avec la rotation actuelle."""
        self.canvas.delete('all')
        
        # Rotation des points
        rotated_a = self._rotate_points(self.a_points, self.logo_angle)
        rotated_bar = self._rotate_points(self.a_bar, self.logo_angle)
        
        # Dessin du A
        self.canvas.create_line(*rotated_a, fill='#8D99AE', width=3)
        self.canvas.create_line(*rotated_bar, fill='#EF233C', width=3)

    def _rotate_points(self, points, angle):
        """Applique une rotation aux points."""
        cx, cy = 50, 50
        cos_a = math.cos(math.radians(angle))
        sin_a = math.sin(math.radians(angle))
        rotated = []
        
        for i in range(0, len(points), 2):
            x, y = points[i] - cx, points[i+1] - cy
            rx = x * cos_a - y * sin_a + cx
            ry = x * sin_a + y * cos_a + cy
            rotated.extend([rx, ry])
        
        return rotated

    def _animate_logo(self):
        """Animation continue du logo."""
        self.logo_angle = (self.logo_angle + 2) % 360
        self._draw_logo()
        self.window.after(50, self._animate_logo)

    def fade_in(self):
        """Animation de fade in."""
        alpha = self.window.attributes('-alpha')
        if alpha < 1.0:
            alpha += 0.1
            self.window.attributes('-alpha', alpha)
            self.window.after(50, self.fade_in)

    def update_progress(self, value: float):
        """Mise à jour de la progression avec animation."""
        self.progress.set_target(value)

    def close(self):
        """Fermeture avec fade out."""
        def fade_out():
            alpha = self.window.attributes('-alpha')
            if alpha > 0:
                alpha -= 0.1
                self.window.attributes('-alpha', alpha)
                self.window.after(50, fade_out)
            else:
                self.window.destroy()
        fade_out()

    def show(self, on_complete: Callable = None):
        """Affiche le splash screen avec toutes les animations."""
        self.window.after(1, self.fade_in)
        self.title_label.start_animation()
        self.status_label.start_animation()
        
        def update_loop(step=0, total_steps=10):
            if step <= total_steps:
                progress = (step / total_steps) * 100
                self.update_progress(progress)
                self.window.after(int(self.duration * 100), 
                                lambda: update_loop(step + 1, total_steps))
            else:
                self.window.after(500, lambda: self.close())
                if on_complete:
                    self.window.after(1000, on_complete)
        
        update_loop()

if __name__ == "__main__":
    def on_splash_complete():
        print("Splash screen terminé, démarrage de l'application...")
    
    splash = SplashScreen(duration=2)
    splash.show(on_splash_complete) 