import datetime as dt
import json
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import core

class App:
    def __init__(self, root):
        self.root=root; root.title('Garmin → Excel Coach'); root.geometry('760x630')
        self.events=queue.Queue()
        frame=ttk.Frame(root,padding=16); frame.pack(fill='both',expand=True)
        ttk.Label(frame,text='Garmin → Excel Coach',font=('Segoe UI',19,'bold')).pack(anchor='w')
        ttk.Label(frame,text='Synchroniser les activités Garmin avec la feuille Semaine, sans modifier les formules.').pack(anchor='w',pady=(0,16))
        self.email=tk.StringVar(); self.password=tk.StringVar(); self.path=tk.StringVar(value=str(Path(__file__).resolve().parent/'Fichier 26 NQ.xlsx'))
        self.week=tk.IntVar(value=1); self.preview=tk.BooleanVar(value=True); self.overwrite=tk.BooleanVar(value=False)
        for label,var,show in [('Email Garmin',self.email,None),('Mot de passe Garmin',self.password,'•')]:
            ttk.Label(frame,text=label).pack(anchor='w'); ttk.Entry(frame,textvariable=var,show=show,width=65).pack(fill='x',pady=(0,9))
        ttk.Label(frame,text='Fichier Excel').pack(anchor='w')
        f=ttk.Frame(frame);f.pack(fill='x',pady=(0,9));ttk.Entry(f,textvariable=self.path).pack(side='left',fill='x',expand=True)
        ttk.Button(f,text='Parcourir',command=self.browse).pack(side='left',padx=(7,0))
        f=ttk.Frame(frame);f.pack(fill='x',pady=(0,9));ttk.Label(f,text='Numéro de semaine (2026) :').pack(side='left')
        ttk.Spinbox(f,from_=1,to=53,textvariable=self.week,width=7).pack(side='left',padx=10)
        ttk.Checkbutton(frame,text='Aperçu uniquement (ne crée pas de fichier Excel)',variable=self.preview).pack(anchor='w')
        ttk.Checkbutton(frame,text='Remplacer les anciennes saisies dans les cellules autorisées',variable=self.overwrite).pack(anchor='w',pady=(0,12))
        self.button=ttk.Button(frame,text='Synchroniser Garmin',command=self.start);self.button.pack(anchor='w',pady=(0,12))
        ttk.Label(frame,text='Journal de synchronisation').pack(anchor='w')
        self.log=tk.Text(frame,height=15,wrap='word');self.log.pack(fill='both',expand=True)
        self.root.after(150,self.poll)
    def browse(self):
        p=filedialog.askopenfilename(filetypes=[('Classeur Excel','*.xlsx')]);
        if p:self.path.set(p)
    def write(self,s):self.events.put(('log',s))
    def start(self):
        if not self.email.get() or not self.password.get():messagebox.showwarning('Connexion','Renseigne ton email et ton mot de passe Garmin.');return
        if not Path(self.path.get()).is_file():messagebox.showerror('Excel','Fichier Excel introuvable.');return
        try: week=int(self.week.get()); assert 1<=week<=53
        except Exception:messagebox.showerror('Semaine','Numéro de semaine invalide.');return
        args=(self.email.get(),self.password.get(),Path(self.path.get()),week,self.preview.get(),self.overwrite.get())
        self.button.configure(state='disabled');self.log.delete('1.0','end');threading.Thread(target=self.worker,args=args,daemon=True).start()
    def worker(self,email,password,src,week,preview,overwrite):
        try:
            from garminconnect import Garmin
            monday=dt.date(2026,1,5)+dt.timedelta(weeks=week-1);sunday=monday+dt.timedelta(days=6)
            self.write(f'Semaine {week} : {monday} → {sunday}\nConnexion Garmin…\n')
            api=Garmin(email,password);api.login();password=None
            raw=api.get_activities_by_date(monday.isoformat(),sunday.isoformat())
            self.write(f'{len(raw)} activité(s) récupérée(s).\n')
            activities=[]
            for a in raw:
                if core.classify(a)=='swim':continue
                self.write(f"• {a.get('activityName','Activité')}\n")
                activities.append(core.activity_info(api,a))
            health={}
            for i in range(7):
                day=monday+dt.timedelta(days=i);health[day.isoformat()]=core.health_day(api,day)
            config={'week_number':week,'week_1_anchor':'AWD14','week_1_monday':'2026-01-05','week_column_stride':33}
            updates,warnings=core.build_updates(activities,health,config)
            preview_path=src.with_name(f'apercu_S{week:02d}.json')
            preview_path.write_text(json.dumps({'cells':{k:{'value':v,'type':t} for k,(v,t) in updates.items()},'warnings':warnings},ensure_ascii=False,indent=2),encoding='utf-8')
            self.write(f'\n{len(updates)} cellules proposées. Aperçu : {preview_path}\n')
            for w in sorted(set(warnings)):self.write('ATTENTION : '+w+'\n')
            if not preview:
                out=src.with_name(f'{src.stem}_rempli_S{week:02d}.xlsx')
                changed,skipped=core.patch_xlsx(src,out,updates,overwrite)
                self.write(f'\n{len(changed)} cellules modifiées ; {len(skipped)} ignorées.\n')
                for cell,reason in skipped:self.write(f'Ignoré {cell}: {reason}\n')
                if changed:self.write(f'Fichier créé : {out}\n')
            else:self.write('\nMode aperçu : classeur non modifié.\n')
            self.events.put(('done','Synchronisation terminée.'))
        except Exception as e:self.events.put(('error',f'{type(e).__name__}: {e}'))
    def poll(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='log':self.log.insert('end',value);self.log.see('end')
                else:
                    self.button.configure(state='normal')
                    (messagebox.showerror if kind=='error' else messagebox.showinfo)('Garmin Excel Coach',value)
        except queue.Empty:pass
        self.root.after(150,self.poll)

if __name__=='__main__':
    root=tk.Tk();App(root);root.mainloop()
