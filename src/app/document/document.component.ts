// word-document.service.ts
import {Component} from '@angular/core';
import {DocumentService} from './word-documen.service';
import {DocumentService2} from './word-documen2.service';
import {DocumentEssaludService} from './word-documen-EsSalud.service';

@Component({
  selector: 'app-document',
  standalone: true,
  templateUrl: './document.component.html',
  styleUrl: './document.component.scss'
})
export class DocumentComponent {
  constructor(private documentService: DocumentService, private documentService2: DocumentService2, public documentEsSaludService: DocumentEssaludService) {
  }

  downloadDocument(): void {
    this.documentService.generateDocument();
  }
  downloadDocument2(): void {
    this.documentService2.generateDocument();
  }


  public documentEsSalud(): void {
    this.documentEsSaludService.generateDocument();
  }
}
