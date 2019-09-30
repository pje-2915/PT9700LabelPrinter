import tkinter as tk
import tkinter.ttk as ttk
import subprocess
import xxsubtype
import shlex
from collections import namedtuple
from subprocess import Popen, PIPE

cmd_start = bytearray.fromhex('1B6961001B40')  # use Esc P, initialise
cmd_landscape = bytearray.fromhex('1B694C00')  # rotated printing off
cmd_format = bytearray.fromhex('1B696D0700')  # minimum margin  1B 69 6D n1 n2
cmd_font = bytearray.fromhex('1B6B')
cmd_font_size = bytearray.fromhex('1B58')
cmd_no_cut = bytearray.fromhex('1B694306')  # chain print + 1/2 cut
cmd_cut = bytearray.fromhex('1B694300')  # full cut
cmd_lf = bytearray.fromhex('0A')
cmd_print = bytearray.fromhex('0C')
font_helsinki = bytearray.fromhex('00')
font_letter_gothic = bytearray.fromhex('01')
font_4_pt = bytearray.fromhex('30')
font_6_pt = bytearray.fromhex('32')
font_9_pt = bytearray.fromhex('33')
font_12_pt = bytearray.fromhex('34')
cmd_set_lf_0 = bytearray.fromhex('1B3300')
cmd_set_lf_24 = bytearray.fromhex('1B3318')
cmd_set_lf_26 = bytearray.fromhex('1B331A')
cmd_line_spacing_00 = bytearray.fromhex('1B3300')
cmd_line_spacing_02 = bytearray.fromhex('1B3302')
cmd_line_spacing_06 = bytearray.fromhex('1B3306')
cmd_line_spacing_24 = bytearray.fromhex('1B3318')
cmd_line_spacing_26 = bytearray.fromhex('1B331A')
cmd_line_spacing_30 = bytearray.fromhex('1B331E')

class Gui(tk.Frame):

    def __init__(self, master):
        self.botanicalName = ""
        self.collectionNumber = ""
        self.notesData = ""
        self.copies = 0

        self.master = master
        ttk.Style().configure("TButton", relief="raised", background="light grey", padding=2)

        tk.Frame.__init__(self, self.master)
        master.config(bg="light grey")
        self.config(bg="light grey")
        self.pack(fill='both', expand=True)
        self.grid(padx=5, pady=5)
        self.name_label = ttk.Label(self, text="Name:")
        self.notes_label = ttk.Label(self, text="Notes / Locality:")
        self.coll_num_label = ttk.Label(self, text="Collection Number:")
        self.count_label = ttk.Label(self, text="Copies:")
        self.name_entry = ttk.Entry(self)
        self.coll_num_entry = ttk.Entry(self)
        self.notes_entry = ttk.Entry(self)

        self.count_entry = ttk.Entry(self)
        reg_fn = self.register(self.is_int)
        self.count_entry.config(validate="key", validatecommand=(reg_fn, '%P'))

        self.print_button = ttk.Button(self, text="Print Label", command=self.print_button_callback)
        self.clear_button = ttk.Button(self, text="Clear", command=self.clear_button_callback)
        self.eject_button = ttk.Button(self, text="Cut & Eject", command=self.cut_eject_button_callback)
        self._separator = ttk.Separator(self, orient="horizontal")
        self.multi_column_listbox = ttk.Treeview(self, show="headings")
        # On Mac can't format this scrollbar - stuck with native styling
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.multi_column_listbox.yview)
        self.multi_column_listbox.config(yscrollcommand=self.scrollbar.set)
        self.queue_button = ttk.Button(self, text="Add to Queue", command=self.queue_button_callback)
        self.print_queue_button = ttk.Button(self, text="Print Queue", command=self.print_queue_button_callback)
        self.clear_queue_button = ttk.Button(self, text="Clear Queue", command=self.clear_queue_button_callback)
        self.edit_queue_entry_button = ttk.Button(self, text="Edit Entry", command=self.edit_queue_entry_callback)
        self.delete_queue_entry_button = ttk.Button(self, text="Delete Entry", command=self.delete_queue_entry_callback)

        self.create_widgets()

    def create_widgets(self):
        self.name_label.config(anchor=tk.E)
        self.name_label.grid(row=0, column=0, sticky=tk.NSEW)
        self.notes_label.config(anchor=tk.E)
        self.notes_label.grid(row=1, column=0, sticky=tk.NSEW)
        self.coll_num_label.config(anchor=tk.E)
        self.coll_num_label.grid(row=2, column=0, sticky=tk.NSEW)
        self.count_label.config(anchor=tk.E)
        self.count_label.grid(row=2, column=2, sticky=tk.NSEW)
        self.name_entry.grid(row=0, column=1, columnspan=3, sticky=tk.EW)
        self.notes_entry.grid(row=1, column=1, columnspan=3, sticky=tk.EW)
        self.coll_num_entry.grid(row=2, column=1)
        self.count_entry.grid(row=2, column=3)
        self.print_button.grid(row=0, column=4, sticky=tk.NSEW)
        self.eject_button.grid(row=1, column=4, sticky=tk.NSEW)
        self.clear_button.grid(row=2, column=4, sticky=tk.NSEW)

        self._separator.grid(row=4, column=0, columnspan=6, sticky=tk.NSEW, pady=10)
        self.queue_button.grid(row=5, column=0, sticky=tk.NSEW)
        self.print_queue_button.grid(row=5, column=1, sticky=tk.NSEW)
        self.clear_queue_button.grid(row=5, column=2, sticky=tk.NSEW)
        self.edit_queue_entry_button.grid(row=5, column=3, sticky=tk.NSEW)
        self.delete_queue_entry_button.grid(row=5, column=4, sticky=tk.NSEW)
        self.multi_column_listbox.grid(row=6, column=0, columnspan=5, sticky=tk.NSEW)
        self.scrollbar.grid(row=6, column=5, sticky=tk.NSEW)

        self.multi_column_listbox["columns"] = ("1", "2", "3", "4")
        self.multi_column_listbox.column("1")
        self.multi_column_listbox.column("2", width=40)
        self.multi_column_listbox.column("3")
        self.multi_column_listbox.column("4", width=20)
        self.multi_column_listbox.heading("0", text="Name")
        self.multi_column_listbox.heading("2", text="Collect Number")
        self.multi_column_listbox.heading("3", text="Notes")
        self.multi_column_listbox.heading("4", text="Copies")

    def is_int(self, inp):
        if inp.isdigit():
            return True
        elif inp is "":
            return True
        else:
            return False

    def get_entry_text(self):
        self.botanicalName = self.name_entry.get()
        self.collectionNumber = self.coll_num_entry.get().upper()
        self.notesData = self.notes_entry.get()
        tmp_copies = self.count_entry.get()
        self.copies = 1

        if len(tmp_copies):
            self.copies = int(self.count_entry.get())

    def clear_entry_text(self):
        self.name_entry.delete(0, "end")
        self.coll_num_entry.delete(0, "end")
        self.notes_entry.delete(0, "end")
        self.count_entry.delete(0, "end")

    def print_button_callback(self):
        self.get_entry_text()
        myLP.print_label(self.botanicalName, self.collectionNumber, self.notesData, self.copies)
        self.clear_entry_text()

    def print_queue_button_callback(self):
        self.get_entry_text()
        myLP.print_all()

    def queue_button_callback(self):
        self.get_entry_text()
        print_queue.append(print_queue_entries(self.botanicalName, self.collectionNumber, self.notesData, self.copies))

        if len(self.botanicalName) > 0:
            self.multi_column_listbox.insert('', 'end', iid=None, values=(self.botanicalName,
                                                                          self.collectionNumber, self.notesData,
                                                                          self.copies))
            self.clear_entry_text()

    def clear_button_callback(self):
        self.clear_entry_text()

    # TODO - are you sure popup
    def clear_queue_button_callback(self):
        print_queue.clear()
        for entry in self.multi_column_listbox.get_children():
            self.multi_column_listbox.delete(entry)

    def cut_eject_button_callback(self):
        eject_cmd_buffer = cmd_start + cmd_cut + cmd_print
        myLP.send_command(eject_cmd_buffer)

    def edit_queue_entry_callback(self):
        focus = self.multi_column_listbox.focus()
        item = self.multi_column_listbox.item(focus)
        self.botanicalName = item['values'][0]
        self.collectionNumber = item['values'][1]
        self.notesData = item['values'][2]
        self.copies = str(item['values'][3])

        self.clear_entry_text()
        self.name_entry.insert(0, self.botanicalName)
        self.coll_num_entry.insert(0, self.collectionNumber)
        self.notes_entry.insert(0, self.notesData)
        self.count_entry.insert(0, self.copies)
        self.delete_queue_entry_callback()

    def delete_queue_entry_callback(self):
        focus = self.multi_column_listbox.focus()
        print("focus is", focus, self.multi_column_listbox.item(focus))
        item = self.multi_column_listbox.item(focus)
        print(item['values'][1])
        self.multi_column_listbox.delete(focus)


class LabelPrinter:

    def __init__(self):
        self.font_ratio = 1.5  # guesswork

    def format_and_print(self, name_data, collection_number, notes_data, copies_count):
        # if there is no notes data then try to split the name

        split_name = name_data
        if len(collection_number) > 0:
            split_name += " ( " + collection_number + " )"

        split_notes = notes_data
        is_split = False
        line_count = 1  # if we don't have a name we wouldn't get here

        if len(notes_data) > 0:
            line_count += 1
            if self.font_ratio * len(notes_data) > len(name_data):
                # notes info is longer to split this rather than the name
                split_notes, is_split = self.line_splitter(notes_data)
            else:
                # leave the notes info as one line and split the name
                split_name, is_split = self.line_splitter(name_data)
        else:
            # we only have the name so may as well split it
            split_name, is_split = self.line_splitter(name_data)
            print("got here", is_split, split_name)

        if is_split:
            line_count += 1

        print("line count", line_count)

        monster_print_buffer = cmd_start + cmd_landscape + cmd_format + cmd_font + font_helsinki + cmd_font_size \
                               + font_12_pt

        ''' 
        Normally the text is better centred if you leave the line feed size alone, but if there's three lines
        then we need to set this so we can fit them all in.
        '''

        if line_count == 3:
            monster_print_buffer += cmd_line_spacing_24

        monster_print_buffer += bytearray(split_name, 'latin-1')

        if len(notes_data) > 0:

            if line_count == 3:
                monster_print_buffer += cmd_line_spacing_26 + cmd_lf + cmd_line_spacing_00 \
                                    + cmd_font_size + font_9_pt \
                                    + bytearray(split_notes, 'latin-1')

            else:
                monster_print_buffer += cmd_lf \
                                        + cmd_font_size + font_9_pt \
                                        + bytearray(split_notes, 'latin-1')

        monster_print_buffer = monster_print_buffer + cmd_no_cut + cmd_print

        for copy in range(copies_count):
            self.send_command(monster_print_buffer)

    def print_label(self, botanical_name, collection_number, notes_data, copies):
        myLP.format_and_print(botanical_name, collection_number, notes_data, copies)

    def print_all(self):
        for queue_entry in print_queue:
            self.print_label(queue_entry.name, queue_entry.collection_num, queue_entry.notes, queue_entry.copies)

    def send_command(self, buffer):
        args = shlex.split("/usr/bin/lpr -l -P Brother_PT_9700PC")
        p = Popen(args, stdin=PIPE)
        p.stdin.write(buffer)
        p.communicate()

    def line_splitter(self, long_line: str):

        # Roughly divide the text in two at a convenient word boundary
        total_len = len(long_line)
        words = long_line.split(' ')
        print("word", words)

        word_index = 0
        len_so_far = 0
        is_split = False

        # Work out which word to split at
        for word in words:
            tmp_length = len(word) + 1  # word length + space (this is rough!)
            if len_so_far + tmp_length > total_len / 2:
                if len_so_far + tmp_length / 2 < total_len / 2:
                    word_index += 1
                break
            len_so_far += tmp_length
            word_index += 1

        print("word index", word_index)
        split_index = word_index

        word_index = 0
        out_string = ""
        for word in words:
            if word_index == split_index and word_index > 0:
                out_string += chr(0x0A)
                is_split = True
            else:
                if word_index > 0:
                    out_string += " "
            out_string += word
            word_index += 1

        return out_string, is_split


if __name__ == '__main__':
    myLP = LabelPrinter()
    root = tk.Tk()
    style = ttk.Style(root)
    style.theme_use('classic')

    root.title('LabelPrinter')
    print_queue = []
    print_queue_entries = namedtuple('PrintQueueEntry', 'name collection_num notes copies')

    myGUI = Gui(root)
    root.mainloop()
